#�ݒ�p�����[�^
$JIRA_URL = "https://niwatk.atlassian.net"
$JIRA_PROJECT = "GT2"
$JIRA_USER = "niwatk@gmail.com"
$JIRA_TOKEN = "token"
$JIRA_ISSUETYPE = @("", "�^�X�N", "�^�X�N", "�^�X�N", "�T�u�^�X�N")
$EXCEL_FILENAME = "JiraExport2.xlsx"
$EXCEL_SHEET = "Sheet1"
$EXCEL_LAST_CELL = 11   #Microsoft���K�肷��R�[�h�l
$EXCEL_START_ROW = 2
$EXCEL_LAST_COLUMN = 7

#Excel�C���X�^���X�𐶐�
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false  # �t�@�C�����J���Ȃ������ꍇ�̃��b�Z�[�W�{�b�N�X�\�����Ȃ��ɂ���

#Excel�t�@�C�����I�[�v��
try {
    $excelFilePath = (Get-ChildItem $EXCEL_FILENAME).FullName
    $book = $excel.Workbooks.Open($excelFilePath, 0, $true)
} catch {
    Write-Error "Excel�t�@�C���I�[�v���Ɏ��s���܂���"
    $excel.Quit()
    exit
}

#�w�薼�̃V�[�g���擾
$sheet = $book.Sheets($EXCEL_SHEET)

#�V�[�g����Z���l��ǂݍ���(�c�FEXCEL_START_ROW�s����f�[�^������ŏI�s�A���F1��ڂ���EXCEL_LAST_COLUMN��܂�)
#2�����z��ł͂Ȃ�1�����z�񂪕Ԃ��Ă���(����ˉE���̏��Ɋi�[)
$xlsData = $sheet.Range($sheet.Cells($EXCEL_START_ROW, 1),
    $sheet.Cells($sheet.Cells.SpecialCells($EXCEL_LAST_CELL).Row, $EXCEL_LAST_COLUMN))

#�V�[�g�̒l���s�P�ʂŘA�z�z��Ɋi�[
$excelIssues = [ordered]@{}
$currentL1Key = $null
$currentL2Key = $null
$currentL3Key = $null
$xlsRowCount = $xlsData.Count / $EXCEL_LAST_COLUMN
for($i = 0; $i -lt $xlsRowCount; $i++) {
    #�e�s�̃Z���l��Ǎ���
    $excelRow = @{
        key = $xlsData[$i * $EXCEL_LAST_COLUMN + 1].Text
        summaryL1 = $xlsData[$i * $EXCEL_LAST_COLUMN + 2].Text
        summaryL2 = $xlsData[$i * $EXCEL_LAST_COLUMN + 3].Text
        summaryL3 = $xlsData[$i * $EXCEL_LAST_COLUMN + 4].Text
        summaryL4 = $xlsData[$i * $EXCEL_LAST_COLUMN + 5].Text
        assignee = $xlsData[$i * $EXCEL_LAST_COLUMN + 6].Text
        duedate = $xlsData[$i * $EXCEL_LAST_COLUMN + 7].Text
        summary = $null
        parent = $null
        hierarchy = 0
    }
    #�L�[�l���󕶎��̏ꍇ(���V�K�ǉ��s)�͉��L�[�l���Z�b�g
    if ($excelRow.key -eq "") {
        $excelRow.key = "*TMP*-" + $i
    }
    #�S���҂��󕶎��̏ꍇ��null�l���Z�b�g
    if ($excelRow.assignee -eq "") {
        $excelRow.assignee = $null
    }
    #���t���󕶎��̏ꍇ��null�l�A�l�������Ă���ꍇ��Jira API�擾���̃t�H�[�}�b�g�ɕϊ�
    if ($excelRow.duedate -eq "") {
        $excelRow.duedate = $null
    } else {
        $excelRow.duedate = ([DateTime]$excelRow.duedate).ToString("yyyy-MM-dd")
    }
    #�K�w���x�����̒l�ݒ�(�K�w���x�����A�v�񖼁A�e�L�[)
    if ($excelRow.summaryL1 -ne "") {
        $excelRow.hierarchy = 1
        $excelRow.summary = $excelRow.summaryL1
        $currentL1Key = $excelRow.key
    } elseif ($excelRow.summaryL2 -ne "") {
        $excelRow.hierarchy = 2
        $excelRow.summary = $excelRow.summaryL2
        $excelRow.parent = $currentL1Key
        $currentL2Key = $excelRow.key
    } elseif ($excelRow.summaryL3 -ne "") {
        $excelRow.hierarchy = 3
        $excelRow.summary = $excelRow.summaryL3
        $excelRow.parent = $currentL2Key
        $currentL3Key = $excelRow.key
    } else {
        $excelRow.hierarchy = 4
        $excelRow.summary = $excelRow.summaryL4
        $excelRow.parent = $currentL3Key
    }

    #Excel�s�ɑΉ������z��A�L�[�l����Excel�s�ւ̘A�z�z��ɂ��ꂼ��ǉ�
    $excelIssues[$excelRow.key] = $excelRow
}

#Excel�t�@�C�������
$book.Close()
Write-Host "Excel�t�@�C���̓Ǎ�����:" $excelFile

#Excel���I��
$excel.Quit()
$excel = $null
Write-Host "Excel�f�[�^�̓Ǎ�����:���� =" $excelIssues.Count

#JIRA API�F�؃w�b�_�̐���
$pair = "${JIRA_USER}:${JIRA_TOKEN}"
$encodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
$headers = @{
    Authorization = "Basic $encodedCredentials"
}

#Jira�ۑ�ꗗ���擾
$startAt = 0
$jiraIssues = [ordered]@{}
while ($true) {
    # �ۑ茟����Jira API�Ăяo�� ������API�͎擾���100��
    $url = "$JIRA_URL/rest/api/3/search?startAt=$startAt&maxResults=100&fields=key,summary,issuetype,parent,duedate,assignee&jql=project=$JIRA_PROJECT ORDER BY key ASC"
    $response = Invoke-RestMethod -Uri $url -Headers $headers
    foreach ($issue in $response.issues) {
        $jiraIssue = @{
            key = $issue.key
            summary = $issue.fields.summary
            duedate = $issue.fields.duedate
            assignee = $issue.fields.assignee.displayName
            parent = $issue.fields.parent.key
            linkparent = $null
            children = @()
        }
        foreach($linkedIssue in $issue.fields.issuelinks) {
            if ($null -ne $linkedIssue.inwardIssue) {
                $jiraIssue.linkparent = $linkedIssue.inwardIssue.key
            }
        }
        $jiraIssues[$issue.key] = $jiraIssue
    }
    if (($response.total - $startAt) -le 100) { #�S���ǂݍ��񂾂烋�[�v�E�o
        break
    }
    $startAt += 100 #���̃y�[�W�ǂݍ��ݗp��startAt�l��100�𑫂�
}
Write-Host "Jira�f�[�^�̓Ǎ�����:���� =" $jiraIssues.Count

#Jira���[�U�ꗗ���擾
$startAt = 0
$jiraUsers = [ordered]@{}
while ($true) {
    # ���[�U������Jira API�Ăяo�� ������API�͎擾���100��
    $url = "$JIRA_URL/rest/api/3/user/assignable/multiProjectSearch?projectKeys=$JIRA_PROJECT&startAt=$startAt&maxResults=100"
    $response = Invoke-RestMethod -Uri $url -Headers $headers
    foreach ($user in $response) {
        $jiraUsers[$user.displayName] = @{
            accountId = $user.accountId
        }
    }
    if (($response.Count - $startAt) -le 100) { #�S���ǂݍ��񂾂烋�[�v�E�o
        break
    }
    $startAt += 100 #���̃y�[�W�ǂݍ��ݗp��startAt�l��100�𑫂�
}
Write-Host "Jira���[�U�̓Ǎ�����:���� =" $jiraUsers.Count

#�V�K�ǉ����X�g�A�X�V���X�g�̏���
$issuesToCreate = @()
$issuesToUpdate = @()

#Jira�ۑ�̐V�K�ǉ����X�g�E�X�V���X�g�̍쐬(Excel��Jira�̍�����r�ɂ��)
foreach ($excelIssue in $excelIssues.Values) {

    if ($excelIssue.key.StartsWith("*TMP*-")) { # �V�K�ǉ��f�[�^�̏ꍇ
        #�V�K�ǉ����X�g�ɒǉ�
        $issuesToCreate += @{
            key = $excelIssue.key
            issuetype = $JIRA_ISSUETYPE[$excelIssue.hierarchy]
            summary = $excelIssue.summary
            assignee = $excelIssue.assignee
            duedate = $excelIssue.duedate
            parent = $excelIssue.parent
            hierarchy = $excelIssue.hierarchy
        }
    } else {  # �����f�[�^�̏ꍇ
        # Jira��̉ۑ�f�[�^���擾
        $jiraIssue = $jiraIssues[$excelIssue.key]
        if ($null -eq $jiraIssue) {
            Write-Host ("Excel��̃L�[[" + $excelIssue.key + "]��Jira��ɂ͑��݂��܂���")
            continue
        }

        #�e�t�B�[���h�̍X�V�L�����u�[���ϐ��ɑ��
        $updated = ($excelIssue.duedate -ne $jiraIssue.duedate)
        $updated = $updated -or ($excelIssue.summary -ne $jiraIssue.summary)
        $updated = $updated -or ($excelIssue.assignee -ne $jiraIssue.assignee)

        #1�ȏ�̃t�B�[���h�ɍX�V����΍X�V���X�g�ɒǉ�
        if ($updated) {
            $issuesToUpdate += @{
                key = $excelIssue.key
                summary = $excelIssue.summary
                assignee = $excelIssue.assignee
                duedate = $excelIssue.duedate
            }
        }
    }
}

# Jira��ɉۑ��ǉ�
$tempKeyToRealkey = @{} # ���L�[�l������L�[�l�ւ̘A�z�z��̃K��������
foreach ($issue in $issuesToCreate) {

    # �e�L�[�l�����L�[�l�̏ꍇ�́A���L�[�l�ɒu������
    if (($null -ne $issue.parent) -and ($issue.parent.StartsWith("*TMP*-"))) {
        $issue.parent = $tempKeyToRealkey[$issue.parent]
    }

    # �S���Ғl�͕\��������ɃA�J�E���gID�l�ɒu������
    $assignee = $null
    if ($null -ne $issue.assignee) {
        $assignee = $jiraUsers[$issue.assignee].accountId
    }

    #�X�V�f�[�^���쐬
    $data = @{
        fields = @{
            project = @{
                key = $JIRA_PROJECT
            }
            issuetype = @{
                name = $issue.issuetype
            }
            summary = $issue.summary
            duedate = $issue.duedate
            parent = @{
                key = $issue.parent
            }
            assignee = @{
                accountId = $assignee
            }
        }
    }

    #�ŉ��w(L4)�ȊO�͐e�q�֌W��IssueLink�ŕ\�����邽��parent�l��null�ɏC��
    if ($issue.hierarchy -ne 4) {
        $data.fields.parent = $null
    }

    #API�p��json�`��(UTF-8)�ɕϊ�
    $body = $data | ConvertTo-Json
    $body = [Text.Encoding]::UTF8.GetBytes($body)

    #Jira API�Ăяo��
    $url = "$JIRA_URL/rest/api/3/issue"
    $apiRes = Invoke-RestMethod -Method POST -Uri $url -Headers $headers -Body $body -ContentType application/json
    Write-Host "Jira�ۑ��ǉ�:key =" $apiRes.key

    #�ǉ����ꂽ�ۑ�ɂ��āA���L�[�ˎ��L�[��Ή����L�^
    $tempKeyToRealkey[$issue.key] = $apiRes.key
    $issue.key = $apiRes.key

    #�ŉ��w��L4�ȊO(=L2, L3)�͐e�q�֌W��IssueLink�ŕ\�����邽�߁AIssueLink��Jira�ɒǉ�
    if (($issue.hierarchy -ne 4) -and ($null -ne $issue.parent)) {
        $data = @{
            inwardIssue = @{
                key = $issue.parent
            }
            outwardIssue = @{
                key = $issue.key
            }
            type = @{
                name = "Hierarchy link (WBSGantt)"
            }
        }
        $body = $data | ConvertTo-Json
        $url = "$JIRA_URL/rest/api/3/issueLink"
        $apiRes = Invoke-RestMethod -Method POST -Uri $url -Headers $headers -Body $body -ContentType application/json
        Write-Host "Jira�ۑ胊���N��ǉ�:key(�e) =" $issue.parent "key(�q) =" $issue.key
    }
}
Write-Host "Jira�f�[�^�̒ǉ�����:���� =" $issuesToCreate.Count

# Jira��̉ۑ���X�V
foreach ($issue in $issuesToUpdate) {
    # �S���Ғl�͕\��������ɃA�J�E���gID�l�ɒu������
    $assignee = $null
    if ($null -ne $issue.assignee) {
        $jiraUser = $jiraUsers[$issue.assignee]
        if ($null -ne $jiraUser) {
            $assignee = $jiraUser.accountId
        } else {
            Write-Host ("�S���Җ�[" + $issue.assignee + "]��Jira��ɑ��݂��܂���")
        }
    }
    
    #�X�V�f�[�^���쐬
    $data = @{
        fields = @{
            summary = $issue.summary
            duedate = $issue.duedate
            assignee = @{
                accountId = $assignee
            }
        }
    }

    #API�p��json�`��(UTF-8)�ɕϊ�
    $body = $data | ConvertTo-Json
    $body = [Text.Encoding]::UTF8.GetBytes($body)

    #Jira API�Ăяo��
    $url = "$JIRA_URL/rest/api/3/issue/" + $issue.key
    $apiRes = Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body -ContentType application/json
    Write-Host "Jira�ۑ���X�V:key =" $issue.key
}
Write-Host "Jira�f�[�^�̍X�V����:���� =" $issuesToUpdate.Count
