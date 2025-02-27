#設定パラメータ
$EXCEL_SEARCH_PATH = ".\*.xlsx"
$EXCEL_SHEET = "Sheet1"
$EXCEL_XLDOWN = -4121 
$EXCEL_START_ROW = 2
$EXCEL_LAST_COLUMN = 3
$EXCEL_ITEMID_COLUMN = 1
$EXCEL_PLAN_COLUMN = 2
$EXCEL_ACTUAL_COLUMN = 3
$JIRA_URL = "https://niwatk.atlassian.net"
$JIRA_PROJECT = "TM"
$JIRA_ISSUE_TYPE = "タスク"
$JIRA_PLAN_FIELD = "customfield_10074"
$JIRA_ACTUAL_FIELD = "customfield_10073"
$JIRA_USER = "xxx@xxx.com"
$JIRA_TOKEN = "token"

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

    #シートからセル値を読み込む(Rangeでデータがある行までを読み込み)
    $xlsData = $sheet.Range($sheet.Cells($EXCEL_START_ROW, 1) 
        ,$sheet.Cells($EXCEL_START_ROW, 1).End($EXCEL_XLDOWN).Offset(0, $EXCEL_LAST_COLUMN - 1))

    #シートの値を行単位で配列変数に格納
    $xlsRowCount = $xlsData.Count / $EXCEL_LAST_COLUMN
    for($i = 0; $i -lt $xlsRowCount; $i++) {
        $excelRowArray += @{
            itemID = $xlsData[$i * $EXCEL_LAST_COLUMN + $EXCEL_ITEMID_COLUMN].Text
            plan = $xlsData[$i * $EXCEL_LAST_COLUMN + $EXCEL_PLAN_COLUMN].Text
            actual = $xlsData[$i * $EXCEL_LAST_COLUMN + $EXCEL_ACTUAL_COLUMN].Text
            file = $file.Substring(0, $file.LastIndexOf('.'))  # 拡張子を除いたファイル名
        }
    }

    #Excelファイルを閉じる
    Write-Host "Excelファイルの読込完了:" $file
    $book.Close()
}

#Excelを終了
$excel.Quit()
$excel = $null
Write-Host "Excelデータの読込完了:件数 =" $excelRowArray.Count

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
    if (($response.total - $startAt) -le 100) { #全件読み込んだらループ脱出
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

    # Jiraの要約名は「Excelファイル名 / 項目ID」の形式とする
    $summary = $excelRow.file + " / " + $excelRow.itemID 

    if ($null -eq $jiraIssues[$summary]) { #Jira上に対象課題が存在しない場合

        #新規追加リストに追加
        $issuesToCreate += @{
            summary = $summary
            plan = $excelPlan
            actual = $excelActual
        }
    } else { #Jira上に対象課題が存在する場合

        # Excelから取得した値を、Jira API対応形式に変換する
        $excelPlan = $null
        if ($excelRow.plan -ne '') {
            $excelPlan = ([DateTime]$excelRow.plan).ToString("yyyy-MM-dd")
        }
        $excelActual = $null
        if ($excelRow.actual -ne '') {
            $excelActual = ([DateTime]$excelRow.actual).ToString("yyyy-MM-dd")
        }

        #各フィールドの更新有無をブール変数に代入
        $planUpdated = ($excelPlan -ne $jiraIssues[$summary].plan)
        $actualUpdated = ($excelActual -ne $jiraIssues[$summary].actual)

        #1つ以上のフィールドに更新あれば更新リストに追加
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
        Write-Host "Jira課題を一括追加:key =" $apiRes.issues[0].key "-" $apiRes.issues[$apiRes.issues.Count-1].key

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

# 削除対象のJira課題抽出(Jira課題一覧からExcelにある要素をすべて削除して、残った物が削除対象)
$issuesToDelete = $jiraIssues.Clone()
foreach($excelRow in $excelRowArray) {
    $issuesToDelete.Remove($excelRow.file + " / " + $excelRow.itemID)
}

# Jira課題の削除
foreach($issue in $issuesToDelete.Values) {
    $url = "$JIRA_URL/rest/api/3/issue/" + $issue.key
    $apiRes = Invoke-RestMethod -Method DELETE -Uri $url -Headers $headers
    Write-Host "Jira課題を削除:key =" $issue.key
}
Write-Host "Jiraデータの削除完了:件数 =" $issuesToDelete.Count
