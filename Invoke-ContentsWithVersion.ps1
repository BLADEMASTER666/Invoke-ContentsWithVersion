<#
    Officeなんかのファイル名の後ろにバージョン番号を追加/更新して起動するスクリプト。
    2019/10/22 K.Takeda

    ファイルバージョンのフォーマットは以下の通り。
    年…4桁(yyyy)、月…2桁(MM)、日…2桁(dd)、リリース番号…2桁(xx)
    "(yyyyMMdd_Rxx)"
#>

param(
    [string]$FilePath
)

# 引数チェック
if($FilePath -eq $null){
    Write-Error -Message "ファイル名が指定されていません"
    exit 8
}

# ファイル存在確認処理
if(-not (Test-Path -Path $FilePath -PathType Leaf)){
    Write-Error -Message '指定されたファイルは存在しません。'
    exit 8
}
$TargetFile = Get-Item -Path $FilePath

# ファイルバージョン文字列処理
$FileOrgName = $TargetFile.BaseName
$FileVerBase = Get-Date -Format "yyyyMMdd"
$FileVerRel  = 0
if($TargetFile.BaseName -match "\(\d{8}_R\d{2}\)$"){
    $FileOrgName = $TargetFile.BaseName.Substring(0,($TargetFile.BaseName.IndexOf($Matches[0])))
    $Matches[0].Substring(11,2)
    [int]$Matches[0].Substring(11,2)
    $FileVerRel  = [int]$Matches[0].Substring(11,2)
    if($Matches[0].Substring(1,8) -match $FileVerBase){
        $FileVerRel += 1
    }else{
        $FileVerRel = 0
    }
}


# 新ファイル名生成
$NewFileName = Join-Path -Path $TargetFile.DirectoryName -ChildPath "$($FileOrgName)($($FileVerBase)_R$('{0:D2}' -f $FileVerRel))$($TargetFile.Extension)"

# ファイルコピー
if((test-path -LiteralPath $NewFileName -PathType Leaf)){
    Write-Error -Message "既にファイルが存在するため、処理を中止します。"
    exit 8
}
Copy-Item -LiteralPath $TargetFile.FullName -Destination $NewFileName

# 対象ファイルを編集
Start-Process -Verb Edit -FilePath $NewFileName
