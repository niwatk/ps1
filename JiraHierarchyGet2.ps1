#�ݒ�p�����[�^
$JIRA_URL = "https://niwatk.atlassian.net"
$JIRA_PROJECT = "GT2"
$JIRA_USER = "niwatk@gmail.com"
$JIRA_TOKEN = "token"
$EXCEL_FILENAME = "JiraExport2.xlsx"

#JIRA API�F�؃w�b�_�̐���
$pair = "${JIRA_USER}:${JIRA_TOKEN}"
$encodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
$headers = @{
    Authorization = "Basic $encodedCredentials"
}

#Jira�ۑ�ꗗ���擾
$startAt = 0                # Jira�f�[�^�擾�̊J�n�ʒu(100�����y�[�W���O�擾)
$jiraIssues = [ordered]@{}  # �擾���ʊi�[�p�̘A�z�z��
while ($true) {
    # �ۑ茟����Jira API�Ăяo�� ������API�͎擾���100��
    $url = "$JIRA_URL/rest/api/3/search?startAt=$startAt&maxResults=100&fields=key,issuelinks,summary,issuetype,parent,duedate,assignee&jql=project=$JIRA_PROJECT ORDER BY key ASC"
    $response = Invoke-RestMethod -Uri $url -Headers $headers
    foreach ($issue in $response.issues) {
        $jiraIssue = @{
            key = $issue.key
            summary = $issue.fields.summary
            duedate = $issue.fields.duedate
            assignee = $issue.fields.assignee.displayName
            issuetype = $issue.fields.issuetype.name
            parent = $issue.fields.parent.key
            children = @()
        }

        # WBS Gantt�̊K�w�\����IssueLink�B���̐e�q�֌W���擾/�ݒ�
        foreach($linkedIssue in $issue.fields.issuelinks) {
            if ($null -ne $linkedIssue.inwardIssue) {
                if (($linkedIssue.type.name -eq "Hierarchy link (WBSGantt)")) {
                    $jiraIssue.parent = $linkedIssue.inwardIssue.key
                }
            }
        }
        $jiraIssues[$issue.key] = $jiraIssue
    }
    if (($response.total - $startAt) -le 100) { #�S���ǂݍ��񂾂烋�[�v�E�o
        break
    }
    $startAt += 100 #���̃y�[�W�ǂݍ��ݗp��startAt�l��100�𑫂�
}
Write-Host ("Jira�f�[�^�̓Ǎ�����:���� = " + $jiraIssues.Count)

# �ǂݍ��񂾉ۑ�S���������āA�e�ˎq�̃����N���`������
$rootIssues = @()
foreach($issue in $jiraIssues.Values) {
    if ($null -eq $issue.parent) {
        $rootIssues += $issue
    } else {
        $jiraIssues[$issue.parent].children += $issue
    }
}

#Excel�C���X�^���X�𐶐�
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

#�V�K�u�b�N�����A������̃V�[�g���擾
$book = $excel.Workbooks.Add()
$sheet = $book.WorkSheets(1)

#�s���FJira�ۑ茏��+1�A�񐔁F7 ��2�����z�������(Excel�ւ�Range�����p)
$rows = $jiraIssues.Count + 1
$xlsData = New-Object "object[,]" $rows, 7

# Excel�̐擪�s(�^�C�g���s)�̒l��z��ɃZ�b�g
$xlsData[0, 0] = "Key"
$xlsData[0, 1] = "L1"
$xlsData[0, 2] = "L2"
$xlsData[0, 3] = "L3"
$xlsData[0, 4] = "L4"
$xlsData[0, 5] = "Assignee"
$xlsData[0, 6] = "Due Date"

# Excel��2�s�ڈȍ~(���f�[�^)�̒l��z��ɃZ�b�g�BJira�c���[�K�w���ォ�瑖������
$i = 1
foreach($issueL1 in $rootIssues) {
    $xlsData[$i, 0] = $issueL1.key
    $xlsData[$i, 1] = $issueL1.summary
    $xlsData[$i, 2] = ""
    $xlsData[$i, 3] = ""
    $xlsData[$i, 4] = ""
    $xlsData[$i, 5] = $issueL1.assignee
    $xlsData[$i, 6] = $issueL1.duedate
    $i++
    foreach($issueL2 in $issueL1.children) {
        # ��2�K�w�̉ۑ�f�[�^�l���Z�b�g
        $xlsData[$i, 0] = $issueL2.key
        $xlsData[$i, 1] = ""
        $xlsData[$i, 2] = $issueL2.summary
        $xlsData[$i, 3] = ""
        $xlsData[$i, 4] = ""
        $xlsData[$i, 5] = $issueL2.assignee
        $xlsData[$i, 6] = $issueL2.duedate
        $i++
        foreach($issueL3 in $issueL2.children) {
            # ��3�K�w�̉ۑ�f�[�^�l���Z�b�g
            $xlsData[$i, 0] = $issueL3.key
            $xlsData[$i, 1] = ""
            $xlsData[$i, 2] = ""
            $xlsData[$i, 3] = $issueL3.summary
            $xlsData[$i, 4] = ""
            $xlsData[$i, 5] = $issueL3.assignee
            $xlsData[$i, 6] = $issueL3.duedate
            $i++
            foreach($issueL4 in $issueL3.children) {
                # ��3�K�w�̉ۑ�f�[�^�l���Z�b�g
                $xlsData[$i, 0] = $issueL4.key
                $xlsData[$i, 1] = ""
                $xlsData[$i, 2] = ""
                $xlsData[$i, 3] = ""
                $xlsData[$i, 4] = $issueL4.summary
                $xlsData[$i, 5] = $issueL4.assignee
                $xlsData[$i, 6] = $issueL4.duedate
                $i++
            }
        }
    }
}

# Excel�̃Z���ɒl���Z�b�g(�������̂��߂�Range�ňꊇ�ł̒l�Z�b�g)
$sheet.Range($sheet.Cells(1, 1),$sheet.Cells($jiraIssues.Count + 1, 7)) = $xlsData

# Excel�̗񕝂𒲐�
$sheet.Columns(1).ColumnWidth = 8  # Key��
$sheet.Columns(2).ColumnWidth = 3  # L1��
$sheet.Columns(3).ColumnWidth = 3  # L2��
$sheet.Columns(4).ColumnWidth = 3  # L3��
$sheet.Columns(5).ColumnWidth = 30 # L4��
$sheet.Columns(6).ColumnWidth = 10 # Assignee��
$sheet.Columns(7).ColumnWidth = 10 # Due Date��

# Excel�t�@�C���ɕۑ�
$excel.DisplayAlerts = $false  # �u�㏑�����܂����v�̕\���Ȃ��ɂ���
$book.SaveAs((Convert-Path .) + "\" + $EXCEL_FILENAME)

#Excel���I��
$excel.Quit()
$excel = $null

Write-Host ("Excel�t�@�C���ւ̃G�N�X�|�[�g����:" + $EXCEL_FILENAME)
