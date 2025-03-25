#設定パラメータ
$JIRA_URL = "https://niwatk.atlassian.net"
$JIRA_PROJECT = "GT2"
$JIRA_USER = "niwatk@gmail.com"
$JIRA_TOKEN = "token"
$JIRA_ISSUETYPE = @("", "タスク", "タスク", "タスク", "サブタスク")
$EXCEL_FILENAME = "JiraExport2.xlsx"
$EXCEL_SHEET = "Sheet1"
$EXCEL_LAST_CELL = 11   #Microsoftが規定するコード値
$EXCEL_START_ROW = 2
$EXCEL_LAST_COLUMN = 7

#Excelインスタンスを生成
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false  # ファイルが開けなかった場合のメッセージボックス表示をなしにする

#Excelファイルをオープン
try {
    $excelFilePath = (Get-ChildItem $EXCEL_FILENAME).FullName
    $book = $excel.Workbooks.Open($excelFilePath, 0, $true)
} catch {
    Write-Error "Excelファイルオープンに失敗しました"
    $excel.Quit()
    exit
}

#指定名のシートを取得
$sheet = $book.Sheets($EXCEL_SHEET)

#シートからセル値を読み込む(縦：EXCEL_START_ROW行からデータがある最終行、横：1列目からEXCEL_LAST_COLUMN列まで)
#2次元配列ではなく1次元配列が返ってくる(左上⇒右下の順に格納)
$xlsData = $sheet.Range($sheet.Cells($EXCEL_START_ROW, 1),
    $sheet.Cells($sheet.Cells.SpecialCells($EXCEL_LAST_CELL).Row, $EXCEL_LAST_COLUMN))

#シートの値を行単位で連想配列に格納
$excelIssues = [ordered]@{}
$currentL1Key = $null
$currentL2Key = $null
$currentL3Key = $null
$xlsRowCount = $xlsData.Count / $EXCEL_LAST_COLUMN
for($i = 0; $i -lt $xlsRowCount; $i++) {
    #各行のセル値を読込み
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
    #キー値が空文字の場合(＝新規追加行)は仮キー値をセット
    if ($excelRow.key -eq "") {
        $excelRow.key = "*TMP*-" + $i
    }
    #担当者が空文字の場合はnull値をセット
    if ($excelRow.assignee -eq "") {
        $excelRow.assignee = $null
    }
    #日付が空文字の場合はnull値、値が入っている場合はJira API取得時のフォーマットに変換
    if ($excelRow.duedate -eq "") {
        $excelRow.duedate = $null
    } else {
        $excelRow.duedate = ([DateTime]$excelRow.duedate).ToString("yyyy-MM-dd")
    }
    #階層レベル毎の値設定(階層レベル数、要約名、親キー)
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

    #Excel行に対応した配列、キー値からExcel行への連想配列にそれぞれ追加
    $excelIssues[$excelRow.key] = $excelRow
}

#Excelファイルを閉じる
$book.Close()
Write-Host "Excelファイルの読込完了:" $excelFile

#Excelを終了
$excel.Quit()
$excel = $null
Write-Host "Excelデータの読込完了:件数 =" $excelIssues.Count

#JIRA API認証ヘッダの生成
$pair = "${JIRA_USER}:${JIRA_TOKEN}"
$encodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
$headers = @{
    Authorization = "Basic $encodedCredentials"
}

#Jira課題一覧を取得
$startAt = 0
$jiraIssues = [ordered]@{}
while ($true) {
    # 課題検索のJira API呼び出し ※このAPIは取得上限100件
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
    if (($response.total - $startAt) -le 100) { #全件読み込んだらループ脱出
        break
    }
    $startAt += 100 #次のページ読み込み用にstartAt値に100を足す
}
Write-Host "Jiraデータの読込完了:件数 =" $jiraIssues.Count

#Jiraユーザ一覧を取得
$startAt = 0
$jiraUsers = [ordered]@{}
while ($true) {
    # ユーザ検索のJira API呼び出し ※このAPIは取得上限100件
    $url = "$JIRA_URL/rest/api/3/user/assignable/multiProjectSearch?projectKeys=$JIRA_PROJECT&startAt=$startAt&maxResults=100"
    $response = Invoke-RestMethod -Uri $url -Headers $headers
    foreach ($user in $response) {
        $jiraUsers[$user.displayName] = @{
            accountId = $user.accountId
        }
    }
    if (($response.Count - $startAt) -le 100) { #全件読み込んだらループ脱出
        break
    }
    $startAt += 100 #次のページ読み込み用にstartAt値に100を足す
}
Write-Host "Jiraユーザの読込完了:件数 =" $jiraUsers.Count

#新規追加リスト、更新リストの準備
$issuesToCreate = @()
$issuesToUpdate = @()

#Jira課題の新規追加リスト・更新リストの作成(ExcelとJiraの差分比較により)
foreach ($excelIssue in $excelIssues.Values) {

    if ($excelIssue.key.StartsWith("*TMP*-")) { # 新規追加データの場合
        #新規追加リストに追加
        $issuesToCreate += @{
            key = $excelIssue.key
            issuetype = $JIRA_ISSUETYPE[$excelIssue.hierarchy]
            summary = $excelIssue.summary
            assignee = $excelIssue.assignee
            duedate = $excelIssue.duedate
            parent = $excelIssue.parent
            hierarchy = $excelIssue.hierarchy
        }
    } else {  # 既存データの場合
        # Jira上の課題データを取得
        $jiraIssue = $jiraIssues[$excelIssue.key]
        if ($null -eq $jiraIssue) {
            Write-Host ("Excel上のキー[" + $excelIssue.key + "]はJira上には存在しません")
            continue
        }

        #各フィールドの更新有無をブール変数に代入
        $updated = ($excelIssue.duedate -ne $jiraIssue.duedate)
        $updated = $updated -or ($excelIssue.summary -ne $jiraIssue.summary)
        $updated = $updated -or ($excelIssue.assignee -ne $jiraIssue.assignee)

        #1つ以上のフィールドに更新あれば更新リストに追加
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

# Jira上に課題を追加
$tempKeyToRealkey = @{} # 仮キー値から実キー値への連想配列のガラを準備
foreach ($issue in $issuesToCreate) {

    # 親キー値が仮キー値の場合は、実キー値に置き換え
    if (($null -ne $issue.parent) -and ($issue.parent.StartsWith("*TMP*-"))) {
        $issue.parent = $tempKeyToRealkey[$issue.parent]
    }

    # 担当者値は表示名を基にアカウントID値に置き換え
    $assignee = $null
    if ($null -ne $issue.assignee) {
        $assignee = $jiraUsers[$issue.assignee].accountId
    }

    #更新データを作成
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

    #最下層(L4)以外は親子関係はIssueLinkで表現するためparent値はnullに修正
    if ($issue.hierarchy -ne 4) {
        $data.fields.parent = $null
    }

    #API用のjson形式(UTF-8)に変換
    $body = $data | ConvertTo-Json
    $body = [Text.Encoding]::UTF8.GetBytes($body)

    #Jira API呼び出し
    $url = "$JIRA_URL/rest/api/3/issue"
    $apiRes = Invoke-RestMethod -Method POST -Uri $url -Headers $headers -Body $body -ContentType application/json
    Write-Host "Jira課題を追加:key =" $apiRes.key

    #追加された課題について、仮キー⇒実キーを対応を記録
    $tempKeyToRealkey[$issue.key] = $apiRes.key
    $issue.key = $apiRes.key

    #最下層のL4以外(=L2, L3)は親子関係はIssueLinkで表現するため、IssueLinkをJiraに追加
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
        Write-Host "Jira課題リンクを追加:key(親) =" $issue.parent "key(子) =" $issue.key
    }
}
Write-Host "Jiraデータの追加完了:件数 =" $issuesToCreate.Count

# Jira上の課題を更新
foreach ($issue in $issuesToUpdate) {
    # 担当者値は表示名を基にアカウントID値に置き換え
    $assignee = $null
    if ($null -ne $issue.assignee) {
        $jiraUser = $jiraUsers[$issue.assignee]
        if ($null -ne $jiraUser) {
            $assignee = $jiraUser.accountId
        } else {
            Write-Host ("担当者名[" + $issue.assignee + "]はJira上に存在しません")
        }
    }
    
    #更新データを作成
    $data = @{
        fields = @{
            summary = $issue.summary
            duedate = $issue.duedate
            assignee = @{
                accountId = $assignee
            }
        }
    }

    #API用のjson形式(UTF-8)に変換
    $body = $data | ConvertTo-Json
    $body = [Text.Encoding]::UTF8.GetBytes($body)

    #Jira API呼び出し
    $url = "$JIRA_URL/rest/api/3/issue/" + $issue.key
    $apiRes = Invoke-RestMethod -Method PUT -Uri $url -Headers $headers -Body $body -ContentType application/json
    Write-Host "Jira課題を更新:key =" $issue.key
}
Write-Host "Jiraデータの更新完了:件数 =" $issuesToUpdate.Count
