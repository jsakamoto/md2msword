# ���̃X�N���v�g�Ɠ����t�H���_�ɂ��� Markdown �t�@�C�� (*.md) �� MS-Word �����ɍ������݂܂��B

# �J�����g�t�H���_���A���̃X�N���v�g�t�@�C��������t�H���_�Ɉړ�����
pushd (split-path $PSCommandPath)

# Markdown > HTML �ϊ��ɁA.NET�p�̃��C�u���� "MarkdownDeep" ��ǂݍ��݁A�C���X�^���X�����B
Add-Type -Path '.\MarkdownDeep.NET.1.5\lib\.NetFramework 3.5\MarkdownDeep.dll'
$mdprocc = New-Object MarkdownDeep.Markdown
$mdprocc.ExtraMode = $true # �\�̃}�[�N�A�b�v�ɑΉ����邽�߂� Markdown Extra ���[�h��L����

# MS-Word �� COM �o�R�ŃC���X�^���X���A�e���v���[�g��Word�����t�@�C����ǂݍ���
# (���̃e���v���[�g�ɁAMarkdown����ϊ����� HTML ���������݂��Ă���)
$msword = New-Object -ComObject "Word.Application"
$msword.Visible = $true
$doc = $msword.Documents.Open(((ls ".\template.docx").FullName))
$doc.Activate()

# HTML ���������ވʒu�ɃJ���b�g���ړ�
# �|�C���g: �������񕶖��܂ňړ����邪�A����1����(1�s)��O�ɃJ���b�g��߂��̂��|�C���g�B
# �����ɃJ���b�g��u���āA������HTML���������ނƁA�t�b�^�[���A��������HTML�ɂ�����̂Ƃ���āA������Ă��܂��B
# (���̃e���v���[�g�ł́A�t�b�^�[�Ƀy�[�W�ԍ��t�B�[���h��u���Ă���A���̐܊p�̃t�b�^�[���������ƌ��Ȃ̂�)
$msword.Selection.EndKey(6) > $null
$msword.Selection.Move(1,-1) > $null

# �J�����g�t�H���_�ɂ���A�g���q=.md �̃t�@�C����񋓂��AMarkdownDeep ���g���� HTML �ɕϊ����AWord�����ɍ������݁B
ls .\*.md | 
sort Name |
% {
    # .md �𕶎���Ƃ��ēǂݍ��� �` HTML������ɕϊ� �` �ꎞ�t�@�C���֏����o��
    $mdcontent = cat $_ -Encoding UTF8 -Raw
    $html = ("<html><body>" + $mdprocc.Transform($mdcontent) + "</body></html>")
    Set-Content -Path .\~temp.html -Encoding UTF8 $html

    # Word�����ցA�ϊ�����(�ꎞ�t�@�C����) HTML ����������
    $msword.Selection.InsertFile(((ls ".\~temp.html").FullName), "", $false, $false, $false)
}

# �ꎞ�t�@�C���𐴑|
del .\~temp.html

# �Ō�ɁAWord�����S�̂�I�����ăt�B�[���h�X�V���邱�ƂŁA�ڎ�������������B
$msword.Selection.WholeStory()
$msword.Selection.Fields.Update() > $null

# ���̃X�N���v�g�ł͂����܂ŁB
# ���D�݂ŁA�ʖ��ŕۑ�������AMS-Word�W���̋@�\�� PDF �փG�N�X�|�[�g������ł��܂��B
