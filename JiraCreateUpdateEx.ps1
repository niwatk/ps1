#設定パラメータ
$EXCEL_SEARCH_PATH = ".\*.xlsx"
$EXCEL_SHEET = "Sheet1"
$EXCEL_START_ROW = 2
$EXCEL_LAST_ROW = 1000
$EXCEL_ITEMID_COLUMN = 1
$EXCEL_PLAN_COLUMN = 2
$EXCEL_ACTUAL_COLUMN = 3
$JIRA_URL = "https://yoursite.atlassian.net"
$JIRA_PROJECT = "TM"
$JIRA_ISSUE_TYPE = "タスク"
$JIRA_PLAN_FIELD = "customfield_10074"
$JIRA_ACTUAL_FIELD = "customfield_10073"
$JIRA_USER = "xxx@xxx.com"
$JIRA_TOKEN = "your_token"

#ディレクトリ内のExcelファイル一覧を取得
$fileList = Get-ChildItem -Path $EXCEL_SEARCH_PATH | Select-Object -ExpandProperty Name
if ($fileList.Count -eq 0) {
    Write-Error "Excelファイルが見つかりませんでした"
    exit
}

#Excelデータ格納用配列を定義
$excelRowArray = @()

#Excelインスタンスを生成
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

#対象となるすべてのExcelファイルからデータを読み込む
foreach($file in $fileList) {

    #Excelファイルをオープン
    try {
        $excelFilePath = (Get-ChildItem $file).FullName
        $book = $excel.Workbooks.Open($excelFilePath, 0, $true)
    } catch {
        Write-Error "Excelファイルオープンに失敗しました"
        $excel.Quit()
        exit
    }

    #指定名のシートを取得
    $sheet = $book.Sheets($EXCEL_SHEET)

    #シートからセル値を読み込む
    for ($i = $EXCEL_START_ROW; $i -le $EXCEL_LAST_ROW; $i++) {
        $excelRow = @{
            itemID = $sheet.Cells.Item($i, $EXCEL_ITEMID_COLUMN).Text
            plan = $sheet.Cells.Item($i, $EXCEL_PLAN_COLUMN).Text
            actual = $sheet.Cells.Item($i, $EXCEL_ACTUAL_COLUMN).Text
            file = $file.Substring(0, $file.LastIndexOf('.'))  # 拡張子を除いたファイル名
        }
        if ($excelRow.itemID -eq "") { # 空行ならループ終了
            break
        } 
        $excelRowArray += $excelRow
    }

    #Excelファイルを閉じる
    Write-Host "Excelデータの読込完了:" $file
    $book.Close()
}

#Excelを終了
$excel.Quit()
$excel = $null

#JIRA API認証ヘッダの生成
$pair = "${JIRA_USER}:${JIRA_TOKEN}"
$encodedCredentials = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
$headers = @{
    Authorization = "Basic $encodedCredentials"
}

#Jira課題一覧を取得
#TODO:要約の重複があったときにエラー or ワーニング出力
$startAt = 0
$jiraIssues = @{}
while ($true) {
    # 課題検索のJira API呼び出し ※このAPIは取得上限100件
    $url = "$JIRA_URL/rest/api/3/search?startAt=$startAt&maxResults=100&jql=project=$JIRA_PROJECT&fields=key,summary,$JIRA_PLAN_FIELD,$JIRA_ACTUAL_FIELD"
    $response = Invoke-RestMethod -Uri $url -Headers $headers
    foreach ($issue in $response.issues) {
        $jiraIssues[$issue.fields.summary] = @{  #要約をキーにした連想配列に読み込み
            key = $issue.key
            plan = $issue.fields.$JIRA_PLAN_FIELD
            actual = $issue.fields.$JIRA_ACTUAL_FIELD
        }
    }
    if ($jiraIssues.Count -eq $response.total) { #全件読み込んだらループ脱出
        break
    } 
    $startAt += 100 #次のページ読み込み用にstartAt値に100を足す
}
Write-Host "Jiraデータの読込完了:件数 =" $jiraIssues.Count

#新規追加リスト、更新リストの準備
$issuesToCreate = @()
$issuesToUpdate = @()

#Jira課題の新規追加リスト・更新リストの作成(ExcelとJiraの差分比較により)
foreach ($excelRow in $excelRowArray) {

    # Excelから取得した値を、Jira API対応形式に変換する
    $summary = $excelRow.file + " / " + $excelRow.itemID # Jiraの要約は「Excelファイル名 / 項目ID」の形式
    $excelPlan = $null
    if ($excelRow.plan -ne '') {
        $excelPlan = ([DateTime]$excelRow.plan).ToString("yyyy-MM-dd")
    }
    $excelActual = $null
    if ($excelRow.actual -ne '') {
        $excelActual = ([DateTime]$excelRow.actual).ToString("yyyy-MM-dd")
    }

    if ($null -eq $jiraIssues[$summary]) { #Jira上に対象課題が存在しない場合

        #新規追加リストに追加
        $issuesToCreate += @{
            summary = $summary
            plan = $excelPlan
            actual = $excelActual
        }
    } else { #Jira上に対象課題が存在する場合

        #各フィールドの更新有無をブール変数に代入
        $planUpdated = ($excelPlan -ne $jiraIssues[$summary].plan)
        $actualUpdated = ($excelActual -ne $jiraIssues[$summary].actual)

        #更新あれば更新リストに追加
        if ($planUpdated -or $actualUpdated) {
            $issuesToUpdate += @{
                key = $jiraIssues[$summary].key
                summary = $summary
                plan = $excelPlan
                actual = $excelActual
            }
        }
    }
}

#新規追加データの入れ物を定義
$data = @{
    issueUpdates = @()
}

# Jiraに課題を新規追加
for ($i = 0; $i -lt $issuesToCreate.Count; $i++){

    # API用の登録データを作成
    $issue = $issuesToCreate[$i]
    $data.issueUpdates += @{
        fields = @{
            summary = $issue.summary
            $JIRA_PLAN_FIELD = $issue.plan
            $JIRA_ACTUAL_FIELD = $issue.actual
            project = @{
                key = $JIRA_PROJECT
            }
            issuetype = @{
                name = $JIRA_ISSUE_TYPE
            }
        }
    }
    # 最後の要素 or 50件ずつの区切りの場合、Jira APIで課題を一括追加
    if (($i -eq $issuesToCreate.Count - 1) -or (($i + 1) % 50 -eq 0)) {
        #API用のjson形式(UTF-8)に変換
        $body = $data | ConvertTo-Json -Depth 5
        $body = [Text.Encoding]::UTF8.GetBytes($body)

        #Jira課題を一括で追加 ※このAPIは上限50件
        $url = "$JIRA_URL/rest/api/3/issue/bulk"
        $apiRes = Invoke-RestMethod -Method POST -Uri $url -Headers $headers -Body $body -ContentType application/json
        Write-Host "Jira課題を一括追加:key =" $apiRes.issues.key[0] "-" $apiRes.issues.key[$apiRes.issues.key.Count-1]

        #データをリセット(空に戻す)
        $data.issueUpdates = @()
    }
}
Write-Host "Jiraデータの追加完了:件数 =" $issuesToCreate.Count


# Jira上の課題を更新
foreach ($issue in $issuesToUpdate) {
    #更新データを作成
    $data = @{
        fields = @{
            $JIRA_PLAN_FIELD = $issue.plan
            $JIRA_ACTUAL_FIELD = $issue.actual
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
