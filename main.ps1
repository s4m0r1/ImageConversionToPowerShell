[void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")
add-type -AssemblyName System.Windows.Forms

# �ݒ�
# �𑜓x�̎w��
$WIDTH  = 1024
$HEIGHT = 576
# �G�N�X�|�[�g��̃t���p�X
# ��: C:\Users\hoge\Documents\ExportSlide\
$EXPORT_PATH = ""

# ���O�ݒ�
$DIR_NAME = [Microsoft.VisualBasic.Interaction]::InputBox("�t�H���_������͂��Ă�������", "�t�H���_������")
$TITLE1   = [Microsoft.VisualBasic.Interaction]::InputBox("1�߂̃^�C�g������͂��Ă�������", "�^�C�g��������")
$TITLE2   = [Microsoft.VisualBasic.Interaction]::InputBox("2�߂̃^�C�g������͂��Ă�������", "�^�C�g��������")
$TITLE3   = [Microsoft.VisualBasic.Interaction]::InputBox("3�߂̃^�C�g������͂��Ă�������", "�^�C�g��������")
$DIR_PATH = $EXPORT_PATH + $DIR_NAME

# �^�C�g���̘A��
# XXXXXX_XXXX_XXX_�@�̌`�ɂ���Ō�͘A�ԂȂ̂ł��Ȃ�
# S�������Ă邩����
$S_EXISTENCE= [System.Windows.Forms.MessageBox]::Show("�^�C�g����S�������Ă��܂����H","�^�C�g��������","YesNo","Question","Button2")
If($S_EXISTENCE -eq "Yes"){
    $TITLE = $TITLE1 + "_" + "S" + "_" + $TITLE2 + "_" + $TITLE3
}Else{
    $TITLE = $TITLE1 + "_" + $TITLE2 + "_" + $TITLE3
}

# �t�H���_�̍쐬
New-Item $DIR_PATH -ItemType Directory

# �t�@�C�����p�J�E���g
$cnt = 1

# �摜�𑜓x�̕ύX
foreach ($arg in $args) {
    $TEMP_CNT = "{0:D2}" -f $cnt
    $TEMP_TITLE = $TITLE + "_" +$TEMP_CNT
    echo $TEMP_TITLE
    mspaint $arg

    # �𑜓x�̕ύX
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

    # �t�@�C���̕ۑ�
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
