[void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")
add-type -AssemblyName System.Windows.Forms

# 設定
# 解像度の指定
$WIDTH  = 1024
$HEIGHT = 576
# エクスポート先のフルパス
# 例: C:\Users\hoge\Documents\ExportSlide\
$EXPORT_PATH = ""

# 名前設定
$DIR_NAME = [Microsoft.VisualBasic.Interaction]::InputBox("フォルダ名を入力してください", "フォルダ名入力")
$TITLE1   = [Microsoft.VisualBasic.Interaction]::InputBox("1つめのタイトルを入力してください", "タイトル名入力")
$TITLE2   = [Microsoft.VisualBasic.Interaction]::InputBox("2つめのタイトルを入力してください", "タイトル名入力")
$TITLE3   = [Microsoft.VisualBasic.Interaction]::InputBox("3つめのタイトルを入力してください", "タイトル名入力")
$DIR_PATH = $EXPORT_PATH + $DIR_NAME

# タイトルの連結
# XXXXXX_XXXX_XXX_　の形にする最後は連番なのでつけない
# Sが入ってるか聞く
$S_EXISTENCE= [System.Windows.Forms.MessageBox]::Show("タイトルにSが入っていますか？","タイトル名入力","YesNo","Question","Button2")
If($S_EXISTENCE -eq "Yes"){
    $TITLE = $TITLE1 + "_" + "S" + "_" + $TITLE2 + "_" + $TITLE3
}Else{
    $TITLE = $TITLE1 + "_" + $TITLE2 + "_" + $TITLE3
}

# フォルダの作成
New-Item $DIR_PATH -ItemType Directory

# ファイル名用カウント
$cnt = 1

# 画像解像度の変更
foreach ($arg in $args) {
    $TEMP_CNT = "{0:D2}" -f $cnt
    $TEMP_TITLE = $TITLE + "_" +$TEMP_CNT
    echo $TEMP_TITLE
    mspaint $arg

    # 解像度の変更
    Start-Sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait("%{H}")
    [System.Windows.Forms.SendKeys]::SendWait("{R}")
    [System.Windows.Forms.SendKeys]::SendWait("{E}")
    [System.Windows.Forms.SendKeys]::SendWait("%{M}")
    Start-Sleep -s 0.5
    [System.Windows.Forms.SendKeys]::SendWait("%{b}")
    [System.Windows.Forms.SendKeys]::SendWait("{RIGHT}")
    [System.Windows.Forms.SendKeys]::SendWait("%{h}")
    [System.Windows.Forms.SendKeys]::SendWait($WIDTH)
    Start-Sleep -s 0.5
    [System.Windows.Forms.SendKeys]::SendWait("%{v}")
    [System.Windows.Forms.SendKeys]::SendWait($HEIGHT)
    [System.Windows.Forms.SendKeys]::SendWait("%{e}")
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -s 0.5

    # ファイルの保存
    [System.Windows.Forms.SendKeys]::SendWait("%{F}")
    [System.Windows.Forms.SendKeys]::SendWait("{V}")
    [System.Windows.Forms.SendKeys]::SendWait("{P}")
    Start-Sleep -s 2
    [System.Windows.Forms.SendKeys]::SendWait("%{N}")
    [System.Windows.Forms.SendKeys]::SendWait($TEMP_TITLE)
    Start-Sleep -s 0.5
    [System.Windows.Forms.SendKeys]::SendWait("^l")
    [System.Windows.Forms.SendKeys]::SendWait("^l")
    Start-Sleep -s 0.5
    [System.Windows.Forms.SendKeys]::SendWait($DIR_PATH)
    Start-Sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait("%{S}")
    Start-Sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -s 1
    [System.Windows.Forms.SendKeys]::SendWait("%{F4}")
    Start-Sleep -s 1
    $cnt += 1
}
