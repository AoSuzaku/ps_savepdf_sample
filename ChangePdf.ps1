##################################################
#
#　Excel⇒PDF変換ツール
#
#　変更履歴
#　　・2020/07/15　新規作成
#
##################################################

# ps1ファイルの格納先を取得
[string]$dir = Split-Path $myInvocation.MyCommand.Path -Parent

# Mainクラス実行
$main = New-Object Main($dir)
$main.ChangePdfExe()

class Main{

    # 変数宣言
    [string]$dir
    [string]$ePath
    [string]$pPath
    [object]$savePdf

    # インスタンス
    Main([string]$inDir){

        $this.dir = $inDir
        $this.ePath = Read-Host "PDF変換を行うExcelファイルの格納先を指定してください。"
        $this.pPath = Read-Host "PDFファイルの出力先を指定してください。"
        $this.savePdf = New-Object SavePdf($this.ePath, $this.pPath, $this.dir)

    }

    # Excel→PDF変換実行
    [void]ChangePdfExe(){

        try{

            # ユーザ設定パスチェック
            if(!$this.savePdf.pathCheck()){
            
                return

            }

            # Excelオブジェクト設定
            $this.savePdf.SetExcelObj()

            # 計測開始
            $this.savePdf.WatchTime($true)

            # ファイル数カウント
            $this.savePdf.CountFile()

            # Excel→PDF変換
            $this.savePdf.ChangePdf()

            # エラー結果出力
            $this.savePdf.OutputErrFile()
        
        }catch{
        
        }finally{

            # 終了処理
            $this.savePdf.WatchTime($false)
            $this.savePdf.ExeEnd()
        
        }
    
    }

}

class SavePdf{

    # 変数宣言
    [int]$totalFile
    [string]$ePath
    [string]$pPath
    [string]$dir
    [object]$excel
    [object]$wb
    [object]$watch
    [object]$time

    # 配列宣言
    [string]$errMsg = @()
    
    # インスタンス
    SavePdf([string]$path1, [string]$path2, [string]$path3){
    
        $this.ePath = $path1
        $this.pPath = $path2
        $this.dir = $path3
        $this.watch = New-Object System.Diagnostics.Stopwatch

    }
    
    # ユーザ設定パスチェック
    [bool]pathCheck(){
    
        # 入力チェック
        if($this.ePath -eq ""){

            echo "Excelファイルの格納先が入力されていません。"
            return $false

        }elseif($this.pPath -eq ""){
        
            echo "PDFファイルの出力先が入力されていません。"
            return $false

        }

        # パス存在チェック
        if(!(Test-Path $this.ePath)){

            echo "Excelファイルの格納先が存在しません。"
            return $false

        }elseif(!(Test-Path $this.pPath)){

            echo "PDFファイルの出力先が存在しません。"
            return $false

        }

        return $true

    }

    # Excelオブジェクト設定
    [void]SetExcelObj(){
    
        $this.excel = New-Object -ComObject Excel.Application
        $this.excel.Visible = $false
        $this.excel.DisplayAlerts = $false
    
    }

    # ファイル数カウント
    [void]CountFile(){
    
        $this.totalFile = (Get-ChildItem $this.ePath -Recurse -Include "*.xls*" -Name | Measure-Object).Count
    
    }

    # Excel→PDF変換
    [void]ChangePdf(){

        Get-ChildItem $this.ePath -Recurse -Include "*.xls*" -Name | % {

            try{

                # 処理カウント
                [int]$cnt += 1
                [string]$status = "{0}／$($this.totalFile)件処理中" -F $cnt
                Write-Progress $status -PercentComplete $cnt -CurrentOperation $currentOperation

                # サブフォルダ配下のパス
                [string]$childPath = $_

                # ファイル　Open
                $this.wb = $($this.excel).Workbooks.Open("$($this.ePath)\$($childPath)", $false, $true, [Type]::Missing, $null)

                # ファイル名（拡張子除く）取得
                [string]$name = [System.IO.Path]::GetFileNameWithoutExtension("$($this.ePath)\$($childPath)")

                $this.wb.ExportAsFixedFormat(0, "$($this.pPath)\$($name).pdf")

                # Close
                $this.wb.Close(0)

            }catch{

                # エラー発生処理
                $this.errMsg += "$($this.ePath)\$($childPath),$($_.Exception.Message)"
        
            }

        }    
    
    }

    # エラー結果出力
    [void]OutputErrFile(){

        try{

            # エラー結果出力
            if($this.errMsg.Length -gt 0){
        
                echo "ファイルパス（絶対パス）,エラーメッセージ" | Out-File -Append "$($this.dir)\errResult.txt"
                echo $this.errMsg | Out-File -Append "$($this.dir)\errResult.txt"

            }

        }catch{
        
            Write-Host "出力処理でエラーが発生しました。"
            Write-Host "ErrMsg：$($_.Exception.Message)"
            throw

        }
    
    }

    # 実行時間計測
    [void]WatchTime([bool]$wFlg){
    
        # 計測開始
        if($wFlg){

            $this.watch.Start()
        
        # 計測終了
        }else{
        
            $this.watch.Stop()
            $this.time = $this.watch.Elapsed
            Write-Host "実行時間："$this.time.TotalSeconds.ToString("0.000")"sec"

        }
    
    }

    # 終了処理
    [void]ExeEnd(){
    
        # メモリ開放
        $this.excel.Quit()
        $this.wb = $null
        $this.time = $null
        $this.excel = $null
        $this.watch = $null

        [GC]::Collect()
    
    }

}