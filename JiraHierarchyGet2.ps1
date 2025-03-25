#設定パラメータ
$JIRA_URL = "https://niwatk.atlassian.net"
$JIRA_PROJECT = "GT2"
$JIRA_USER = "niwatk@gmail.com"
$JIRA_TOKEN = "token"
$EXCEL_FILENAME = "JiraExport2.xlsx"

#JIRA API認証ヘッダの生成
$pair = "${JIRA_USER}:${JIRA_TOKEN}"
$encodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
$headers = @{
    Authorization = "Basic $encodedCredentials"
}

#Jira課題一覧を取得
$startAt = 0                # Jiraデータ取得の開始位置(100件ずつページング取得)
$jiraIssues = [ordered]@{}  # 取得結果格納用の連想配列
while ($true) {
    # 課題検索のJira API呼び出し ※このAPIは取得上限100件
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

        # WBS Ganttの階層表現はIssueLink。その親子関係を取得/設定
        foreach($linkedIssue in $issue.fields.issuelinks) {
            if ($null -ne $linkedIssue.inwardIssue) {
                if (($linkedIssue.type.name -eq "Hierarchy link (WBSGantt)")) {
                    $jiraIssue.parent = $linkedIssue.inwardIssue.key
                }
            }
        }
        $jiraIssues[$issue.key] = $jiraIssue
    }
    if (($response.total - $startAt) -le 100) { #全件読み込んだらループ脱出
        break
    }
    $startAt += 100 #次のページ読み込み用にstartAt値に100を足す
}
Write-Host ("Jiraデータの読込完了:件数 = " + $jiraIssues.Count)

# 読み込んだ課題全件走査して、親⇒子のリンクを形成する
$rootIssues = @()
foreach($issue in $jiraIssues.Values) {
    if ($null -eq $issue.parent) {
        $rootIssues += $issue
    } else {
        $jiraIssues[$issue.parent].children += $issue
    }
}

#Excelインスタンスを生成
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

#新規ブックを作り、書込先のシートを取得
$book = $excel.Workbooks.Add()
$sheet = $book.WorkSheets(1)

#行数：Jira課題件数+1、列数：7 の2次元配列を準備(ExcelへのRange書込用)
$rows = $jiraIssues.Count + 1
$xlsData = New-Object "object[,]" $rows, 7

# Excelの先頭行(タイトル行)の値を配列にセット
$xlsData[0, 0] = "Key"
$xlsData[0, 1] = "L1"
$xlsData[0, 2] = "L2"
$xlsData[0, 3] = "L3"
$xlsData[0, 4] = "L4"
$xlsData[0, 5] = "Assignee"
$xlsData[0, 6] = "Due Date"

# Excelの2行目以降(実データ)の値を配列にセット。Jiraツリー階層を上から走査する
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
        # 第2階層の課題データ値をセット
        $xlsData[$i, 0] = $issueL2.key
        $xlsData[$i, 1] = ""
        $xlsData[$i, 2] = $issueL2.summary
        $xlsData[$i, 3] = ""
        $xlsData[$i, 4] = ""
        $xlsData[$i, 5] = $issueL2.assignee
        $xlsData[$i, 6] = $issueL2.duedate
        $i++
        foreach($issueL3 in $issueL2.children) {
            # 第3階層の課題データ値をセット
            $xlsData[$i, 0] = $issueL3.key
            $xlsData[$i, 1] = ""
            $xlsData[$i, 2] = ""
            $xlsData[$i, 3] = $issueL3.summary
            $xlsData[$i, 4] = ""
            $xlsData[$i, 5] = $issueL3.assignee
            $xlsData[$i, 6] = $issueL3.duedate
            $i++
            foreach($issueL4 in $issueL3.children) {
                # 第3階層の課題データ値をセット
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

# Excelのセルに値をセット(高速化のためにRangeで一括での値セット)
$sheet.Range($sheet.Cells(1, 1),$sheet.Cells($jiraIssues.Count + 1, 7)) = $xlsData

# Excelの列幅を調整
$sheet.Columns(1).ColumnWidth = 8  # Key列幅
$sheet.Columns(2).ColumnWidth = 3  # L1列幅
$sheet.Columns(3).ColumnWidth = 3  # L2列幅
$sheet.Columns(4).ColumnWidth = 3  # L3列幅
$sheet.Columns(5).ColumnWidth = 30 # L4列幅
$sheet.Columns(6).ColumnWidth = 10 # Assignee列幅
$sheet.Columns(7).ColumnWidth = 10 # Due Date列幅

# Excelファイルに保存
$excel.DisplayAlerts = $false  # 「上書きしますか」の表示なしにする
$book.SaveAs((Convert-Path .) + "\" + $EXCEL_FILENAME)

#Excelを終了
$excel.Quit()
$excel = $null

Write-Host ("Excelファイルへのエクスポート完了:" + $EXCEL_FILENAME)
