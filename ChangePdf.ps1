######################################################################
#
#　Excel⇒PDF変換ツール
#
#　変更履歴
#　　・2020/07/15　新規作成
#　　・2020/07/21　クラス：PrntrCnf（プリンター情報設定／取得）を追加
#
######################################################################

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
    [object]$prntrCnf

    # コンストラクタ
    Main([string]$inDir){

        $this.dir = $inDir
        $this.ePath = Read-Host "PDF変換を行うExcelファイルの格納先を指定してください。"
        $this.pPath = Read-Host "PDFファイルの出力先を指定してください。"
        $this.savePdf = New-Object SavePdf($this.ePath, $this.pPath, $this.dir)
        $this.prntrCnf = New-Object PrntrCnf

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

            # プリンター設定情報取得
            $this.prntrCnf.GetPrntr()
            
            # プリンター設定情報変更
            $this.prntrCnf.SetPrntr("Microsoft XPS Document Writer")

            # Excel→PDF変換
            $this.savePdf.ChangePdf()

            # プリンター設定情報変更
            $this.prntrCnf.SetPrntr($this.prntrCnf.default)
            $this.prntrCnf.ResetObj()

            # エラー結果出力
            $this.savePdf.OutputErrFile()

            # 終了処理
            $this.savePdf.WatchTime($false)
            $this.savePdf.ExeEnd()
        
        }catch{
        
            Write-Host "処理中にエラーが発生しました。"
        
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
    
    # コンストラクタ
    SavePdf([string]$path1, [string]$path2, [string]$path3){
    
        $this.ePath = $path1
        $this.pPath = $path2
        $this.dir = $path3

    }
    
    # ユーザ設定パスチェック
    [bool]pathCheck(){
    
        # 入力チェック
        if($this.ePath -eq ""){

            Write-Host "Excelファイルの格納先が入力されていません。"
            return $false

        }elseif($this.pPath -eq ""){
        
            Write-Host "PDFファイルの出力先が入力されていません。"
            return $false

        }

        # パス存在チェック
        if(!(Test-Path $this.ePath)){

            Write-Host "Excelファイルの格納先が存在しません。"
            return $false

        }elseif(!(Test-Path $this.pPath)){

            Write-Host "PDFファイルの出力先が存在しません。"
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
                Write-Progress $status -PercentComplete ($cnt/$this.totalFile*100) -CurrentOperation $currentOperation

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

            $this.watch = New-Object System.Diagnostics.Stopwatch
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

class PrntrCnf{

	# 変数宣言
	[string]$default
	[object]$pd

	# コンストラクタ
	PrntrCnf(){
	
		Add-Type -Assembly System.Drawing
		$this.pd = New-Object System.Drawing.Printing.PrintDocument

	}

	# プリンター設定取得
	[void]GetPrntr(){

		$this.default = $this.pd.Name

	}
	
	# プリンター設定
	[void]SetPrntr([string]$prntr){
	
		(Get-WmiObject -ComputerName . -Class Win32_Printer -Filter "Name='$($prntr)'").SetDefaultPrinter()
	
	}
	
	# オブジェクト初期化
	[void]ResetObj(){
	
		$this.pd = $null
	
	}

}