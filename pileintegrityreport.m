function [  ] = pileintegrityreport( varargin )
%   PileIntegrityReport  ����׮�������Գ�����͸�䷨��ⱨ�� v20200817
%   ����ݱ��������һ���ļ���
%   Ŀ¼�±����С���ⱨ�桱�ļ��У��������������ļ�����������������ļ�����
%   pileintegrityreport( 2 )   ������1������������ʾ�������š�2�ļ�ⱨ��
%   pileintegrityreport( 2, 8 )   ������2������������ʾ�������š�2~8�ļ�ⱨ��
%   v20200817  ������õ��ļ��Ƿ���ڵ��ж�


switch nargin
    case 0                                                                  % ���������ʱִ����������
        beginNo1 = input('���������ʼ��� : ' );    %   �������뿪ʼ���
        endNo1 = input('�������������� : ' );      %   ��������������
    case 1                                                                  %  ����1������ʱִ������������һ������
        beginNo1 = varargin{1};
        endNo1 =  varargin{1};
    case 2                                                                  %   ����2������ʱִ����������
        beginNo1 = varargin{1};
        endNo1 =  varargin{2};
    otherwise
        error('More Than 1 Pile Parameters File !');
end

tic;                                                                        %  ��ʼ��ʱ

%%  �������

disp('�������ݲ�����Ϣ');

[~,~,raw1] = xlsread('������Ϣ.xlsx');                                       %  �����ļ�
[~,~,raw2] = xlsread('�����Ϣ.xlsx');                                       %  �����ļ�

scgsytspec_3 = [ pwd '\�����������ļ�\3������ܷ�λʾ��ͼ.jpg' ];              %  3������ܷ�λʾ��ͼ���·��
scgsytspec_4 = [ pwd '\�����������ļ�\4������ܷ�λʾ��ͼ.jpg' ];              %  4������ܷ�λʾ��ͼ���·��
logospc = [ pwd '\�����������ļ�\����ʡ��Ժ���̼��Ƽ����޹�˾�ձ�.jpg' ];     %  ��˾logo
filedirspc = [pwd '\��ⱨ��'];                                                 %  ��ⱨ�����ļ���

if ~exist(scgsytspec_3, 'file')
    error(['�ļ������ڣ�', scgsytspec_3]);
end
if ~exist(scgsytspec_3, 'file')
    error(['�ļ������ڣ�', scgsytspec_4]);
end
if ~exist(scgsytspec_3, 'file')
    error(['�ļ������ڣ�', logospc]);
end
if ~exist(filedirspc, 'dir')
    mkdir(filedirspc);
end

%%  ����word�ĵ��Ļ���������Ϣ

logosize = [ 209.25 , 37.65 ];                                              %  ����logo�ߴ磬[ �� �� ��(��) ]
exportAsPDF = 1 ;                                                           %  �Ƿ����PDF��ʽ�ļ�   1 - �ǣ�0 - ��
exportFlyPage = 1 ;                                                         %  �Ƿ������ҳ��ע�����   1 - �ǣ�0 - ��
supervisionPattern = 1 ;                                                    %  ����ģʽѡ�  1 - �ܼ�졢פ�ؼ���ģʽ�� 2 - ��һ����λģʽ

%%  ԭʼ���ݴ���


[ bgxh , bgbh , jgwmc , dtbh , jzbh , concreteStrength , sjzc , sczc , sjzj , gzfl , gzrq , jcrq , jcjg , ...              %    ����������������庬���readPara
    qxfx , cggs , cgjj1_2 , cgjj1_3 , cgjj2_3 , cgjj1_4 , cgjj2_4 , cgjj3_4 , fjys , ...
    cdjj , yqsb , jcbz , jzbg , gcmc , wtdw, jcnr, jclb , bxry , jcry , fhry , ...
    shry , pzry , pzjb , jcdw , dwdz , dwdh , yzbm , dwcz , bgrq , bgys , ...
    jsdw , kcdw , sjdw , zjb , jldw , zdjl , sgdw , jcxs , sbjgxs ,  ...
    jcxm , jcff , jcsl , gcgk , dztj , sggc , jcgc , zwpmbzt , bz , ...
    zysx , sjzTbg , sjzBbg , sczTbg , sczBbg ] ...
    = readParameters( raw1 , raw2 );


beginNo = beginNo1 - str2double(bgxh{1}) + 1 ;                              % ����������Ϊ��1��ͷ������
endNo = endNo1 - str2double(bgxh{1}) + 1 ;                                  % ����������Ϊ��1��ͷ������

zwpmbztspec = fullfile(pwd,'\�����������ļ�\',zwpmbzt);
for ii = beginNo:1:endNo  
    if ~exist(string(zwpmbztspec{ii}), 'file')
        error(['�ļ������ڣ�', zwpmbztspec{ii}]);
    end
end

%% ���������ҳ���

disp('������������ҳ')
cover( beginNo , endNo , logospc , logosize , exportAsPDF , exportFlyPage ,  ...
    zysx , bgxh , bgbh , gcmc , jcxm , jcff , jcnr , jcdw , bgrq )                           %  ���������ҳ�������
disp('���������ҳ������');

%%

disp('�������ҳ')
reportPage( beginNo , endNo , scgsytspec_3 , scgsytspec_4 , exportAsPDF , supervisionPattern , ...                      %   ����ҳ�������
    bgxh , bgbh , jgwmc , dtbh , jzbh , concreteStrength , sjzc , sczc , sjzj , gzfl , gzrq , jcrq , jcjg , ...
    qxfx , cggs , cgjj1_2 , cgjj1_3 , cgjj2_3 , cgjj1_4 , cgjj2_4 , cgjj3_4 , fjys , ...
    cdjj , yqsb , jcbz , jzbg , gcmc , wtdw , jcnr , jclb , bxry , jcry , fhry , ...
    shry , pzry , pzjb , jcdw , dwdz , dwdh , yzbm , dwcz , bgrq , bgys , ...
    jsdw , kcdw , sjdw , zjb , jldw , zdjl , sgdw , jcxs , sbjgxs ,  ...
    jcxm , jcff , jcsl , gcgk , dztj , sggc , jcgc , zwpmbzt , bz , ...
    zysx , sjzTbg , sjzBbg , sczTbg , sczBbg );
disp('����ҳ������');


%%

disp('����������');
toc;                                                                       %  ������ʱ
disp(' ');

% reply = input('Do you want more? Y/N [Y]: ', 's');
% if isempty(reply)
%     reply = 'Y';
% end

end


function [] = cover( beginNo , endNo , logospc , logosize , exportAsPDF , exportFlyPage , ...
    zysx , bgxh , bgbh , gcmc , jcxm , jcff , jcnr , jcdw , bgrq )
%  cover  ����������YJS-150  v20171122


%%  ����word�ĵ��Ļ���������Ϣ

paperSize = 'wdPaperA4';                                                    %  ����ֽ�Ŵ�С    'wdPaperA4'     'wdPaperA3'   'wdPaperB5' ......
orientation = 'wdOrientPortrait';                                           %  ����ֽ�ŷ���   'wdOrientPortrait'  'wdOrientLandscape'
margin = [42.55, 42.55, 56.7, 56.7];                                      %  ����ҳ�߾࣬��λpoint(��)    [TopMargin, BottomMargin, LeftMargin, RightMargin]
headerFooterDistance = [42.55, 42.55];                                       %  ����ҳüҳ�ž�߽磬��λpoint(��)   [ҳü��ҳ��]


%%  ����Microsoft Word ������

disp('    ���� Microsoft Word ������');

try                                                                         %  ����actxserver��������Microsoft Word ������
    word = actxGetRunningServer('Word.Application');                        %  ��word�������Ѿ��򿪣���������Word
catch
    word = actxserver('Word.Application');                                  %  ��word������û�д򿪣�����һ�������ؾ��Word
end

set(word,'Visible',0);                                                      %  ���÷����������Ϊ�ɼ�״̬, 1 - �ɼ��� 0 - ���ɼ�
document = invoke(word.Documents,'Add');                                    %  ��һ���հ��ĵ�

selection = word.Selection;                                                 %  ����Word.Selection�ӿھ��
leftTabstop = margin(3) + 135;                                              %  ����ҳü�Ҷ˶����Ʊ��λ��

%%  ҳ������

disp('    �ĵ�ҳ������');
document.PageSetup.PaperSize = paperSize;                                   % ����ֽ�Ŵ�С��A4��A3��B5
document.PageSetup.Orientation = orientation;                               % ����ֽ�ŷ���
document.PageSetup.TopMargin = margin(1);                                   % ҳ��߾�����
document.PageSetup.BottomMargin = margin(2);
document.PageSetup.LeftMargin = margin(3);
document.PageSetup.RightMargin = margin(4);
document.PageSetup.HeaderDistance = headerFooterDistance(1);                %  ����ҳü��߽����
document.PageSetup.FooterDistance = headerFooterDistance(2);                %  ����ҳ�ž�߽����

%% ��ʽ����

disp('    �ĵ���ʽ����');


try
    userStyles = document.Styles.Item('����');
catch
    userStyles = document.Styles.Add('����');
end                                                                         %  ���á����ġ���ʽ
userStyles.Font.Name = '����_GB2312';                                       %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����_GB2312';                                %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 14.0;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';           %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)  ��ֵ��������
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  �����о�    'wdLineSpaceSingle' 'wdLineSpace1pt5'

%%  ������������ҳ

for ii = beginNo : 1 : endNo
    
    disp(['    ������ ', bgxh{ii} ,' ������( ' , bgbh{ii} ,' )����']);
    
    if ii ~= beginNo
        selection.InsertBreak(2);                                           %  �����ҳ��0��ֽڷ�2
    end
    
    selection.Style = '����';
    selection.ParagraphFormat.SpaceAfter = 42;                              %  �κ���(��)
    selection.TypeParagraph;
    selection.Text = strcat('�����ţ�', bgbh{ii} );                        %  �ڶ��У�������
    selection.Font.Name = '������';                                          %  ����Ĭ������Ϊ����
    selection.Font.NameFarEast = '������';                                   %  ������������Ϊ����
    selection.Font.NameAscii = 'Times New Roman';                            %  ����Ascii����ΪTimes New Roman
    selection.Font.NameOther = 'Times New Roman';                            %  ���������ַ�����ΪTimes New Roman
    selection.Font.Size = 12.0;                                              %  ���������С(��)
    selection.Font.Bold = 1;                                                 %  ��������Ӵ�
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphRight';           %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
    selection.ParagraphFormat.SpaceBefore = 6;                               %  ��ǰ���(��)
    selection.ParagraphFormat.SpaceAfter = 12;                               %  �κ���(��)
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.Text = gcmc{ii};                                               %
    selection.Font.Name = '��������';                                         %  ����Ĭ������Ϊ����
    selection.Font.NameFarEast = '��������';                                  %  ������������Ϊ����
    selection.Font.NameAscii = 'Times New Roman';                             %  ����Ascii����ΪTimes New Roman
    selection.Font.NameOther = 'Times New Roman';                             %  ���������ַ�����ΪTimes New Roman
    selection.Font.Size = 18.0;                                               %  ���������С(��)
    selection.Font.Bold = 1;                                                  %  ��������Ӵ�
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';           %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
    selection.ParagraphFormat.SpaceBefore = 0;                                %  ��ǰ���(��)
    selection.ParagraphFormat.SpaceAfter = 0;                                 %  �κ���(��)
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.Font.Size = 18.0;                                               %  ���������С(��)
    
    selection.TypeParagraph;
    selection.Text = strcat( '��������ע׮׮��������',  11 , '������͸�䷨��ⱨ��');                %  �м�������   selection.Text = strcat( '��������ע׮׮��������',  11 , '������͸�䷨��ⱨ��');    selection.Text = strcat( jcxm{ii}, jcff{ii}, '��ⱨ��');
    selection.Font.Name = '��������';                                        %  ����Ĭ������Ϊ����
    selection.Font.NameFarEast = '��������';                                 %  ������������Ϊ����
    selection.Font.NameAscii = 'Times New Roman';                            %  ����Ascii����ΪTimes New Roman
    selection.Font.NameOther = 'Times New Roman';                            %  ���������ַ�����ΪTimes New Roman
    selection.Font.Size = 36.0;                                              %  ���������С(��)
    selection.Font.Bold = 1;                                                 %  ��������Ӵ�
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';          %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.ParagraphFormat.SpaceAfter = 222;                              %  �κ���(��)
    
    selection.TypeParagraph;                                                 %  �س�������һ��
    selection.Style = '����';
    handle1 = selection.InlineShape.AddPicture(logospc);                     % ���빫˾logo
    handle1.Width = logosize(1);                                             % logo����
    handle1.Height = logosize(2);
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.Text = jcdw{ii};
    selection.Font.Name = '����_GB2312';                                      %  ����Ĭ������Ϊ����
    selection.Font.NameFarEast = '����_GB2312';                               %  ������������Ϊ����
    selection.Font.NameAscii = 'Times New Roman';                             %  ����Ascii����ΪTimes New Roman
    selection.Font.NameOther = 'Times New Roman';                             %  ���������ַ�����ΪTimes New Roman
    selection.Font.Size = 18.0;                                               %  ���������С(��)
    selection.Font.Bold = 0;                                                  %  ��������Ӵ�
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';           %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
    selection.MoveDown;
    
    selection.TypeParagraph;
%     yyyy = num2str( year( bgrq{ii} ) );                                  %  ��ȡ������
%     mm = num2str( month( bgrq{ii} ) );
%     dd = num2str( day( bgrq{ii} ) );
    yyyy = num2str( datestr( strtok(bgrq{ii}) , 10 ) );                            %  ����datestr������ȡ������
    mm = num2str( datestr( strtok(bgrq{ii}) , 5 ) );
    dd = num2str( datestr( strtok(bgrq{ii}) , 7 ) );
    
    yyyy = strrep(yyyy, '0', '��');
    yyyy = strrep(yyyy, '1', 'һ');
    yyyy = strrep(yyyy, '2', '��');
    yyyy = strrep(yyyy, '3', '��');
    yyyy = strrep(yyyy, '4', '��');
    yyyy = strrep(yyyy, '5', '��');
    yyyy = strrep(yyyy, '6', '��');
    yyyy = strrep(yyyy, '7', '��');
    yyyy = strrep(yyyy, '8', '��');
    yyyy = strrep(yyyy, '9', '��');
  
    mm = strrep(mm, '10', 'ʮ');
    mm = strrep(mm, '11', 'ʮһ');
    mm = strrep(mm, '12', 'ʮ��');
    mm = strrep(mm, '01', 'һ');
    mm = strrep(mm, '02', '��');
    mm = strrep(mm, '03', '��');
    mm = strrep(mm, '04', '��');
    mm = strrep(mm, '05', '��');
    mm = strrep(mm, '06', '��');
    mm = strrep(mm, '07', '��');
    mm = strrep(mm, '08', '��');
    mm = strrep(mm, '09', '��');

    dd = strrep(dd, '0', '��');
    dd = strrep(dd, '1', 'һ');
    dd = strrep(dd, '2', '��');
    dd = strrep(dd, '3', '��');
    dd = strrep(dd, '4', '��');
    dd = strrep(dd, '5', '��');
    dd = strrep(dd, '6', '��');
    dd = strrep(dd, '7', '��');
    dd = strrep(dd, '8', '��');
    dd = strrep(dd, '9', '��');
    selection.Text = strcat( yyyy , '��' , mm , '��');
    selection.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
    selection.Font.NameFarEast = '����';                                       %  ������������Ϊ����
    selection.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
    selection.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
    selection.Font.Size = 18.0;                                                %  ���������С(��)
    selection.Font.Bold = 0;                                                   %  ��������Ӵ�
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.TypeParagraph;
    selection.Text = '��YJS-1150-F/0';
    selection.Font.Name = '������';                                            %  ����Ĭ������Ϊ����
    selection.Font.NameFarEast = '������';                                     %  ������������Ϊ����
    selection.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
    selection.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
    selection.Font.Size = 10.5;                                                %  ���������С(��)
    selection.Font.Bold = 0;                                                   %  ��������Ӵ�
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';              %  ���ö��뷽ʽ 'wdAlignParagraph
    selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
    selection.MoveDown;
    
    if exportFlyPage == 1
        disp(['    ������ ', bgxh{ii} ,' ������( ' , bgbh{ii} ,' )��ҳ��ע�����']);
        selection.InsertBreak(0);                                             %  �����ҳ��0��ֽڷ�2
        selection.Style = '����';
        selection.ParagraphFormat.SpaceAfter = 28;                              %  �κ���(��)
        selection.TypeParagraph;
        selection.Text = 'ע �� �� ��';                                       %  ����
        selection.Font.Name = '����';                                         %  ����Ĭ������Ϊ����
        selection.Font.NameFarEast = '����';                                  %  ������������Ϊ����
        selection.Font.NameAscii = 'Times New Roman';                         %  ����Ascii����ΪTimes New Roman
        selection.Font.NameOther = 'Times New Roman';                         %  ���������ַ�����ΪTimes New Roman
        selection.Font.Size = 28.0;                                           %  ���������С(��)
        selection.Font.Bold = 1;                                              %  ��������Ӵ�
        selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';       %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
        selection.ParagraphFormat.SpaceBefore = 0;                            %  ��ǰ���(��)
        selection.ParagraphFormat.SpaceAfter = 17;                            %  �κ���(��)
        selection.MoveDown;
        
        selection.TypeParagraph;
        selection.Text = zysx{ii};                                            %  "ע������"����
        selection.Font.Name = '����_GB2312';                                  %  ����Ĭ������Ϊ����
        selection.Font.NameFarEast = '����_GB2312';                           %  ������������Ϊ����
        selection.Font.NameAscii = 'Times New Roman';                         %  ����Ascii����ΪTimes New Roman
        selection.Font.NameOther = 'Times New Roman';                         %  ���������ַ�����ΪTimes New Roman
        selection.Font.Size = 14.0;                                           %  ���������С(��)
        selection.Font.Bold = 0;                                              %  ��������Ӵ�
        selection.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';      %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
        selection.ParagraphFormat.CharacterUnitFirstLineIndent = -2;          %  ��������(�ַ�)����Ϊ������������Ϊ��������
        selection.ParagraphFormat.SpaceBefore = 0;                            %  ��ǰ���(��)
        selection.ParagraphFormat.SpaceAfter = 0;                             %  �κ���(��)
        selection.MoveDown;
        
        selection.TypeParagraph;
        selection.ParagraphFormat.SpaceAfter = 42;                                  %  �κ���(��)
        
        selection.TypeParagraph;
        selection.Text = strcat( 09 , 09 , jcdw{ii});
        selection.Font.Name = '����_GB2312';                                   %  ����Ĭ������Ϊ����
        selection.Font.NameFarEast = '����_GB2312';                            %  ������������Ϊ����
        selection.Font.NameAscii = 'Times New Roman';                          %  ����Ascii����ΪTimes New Roman
        selection.Font.NameOther = 'Times New Roman';                          %  ���������ַ�����ΪTimes New Roman
        selection.Font.Size = 14.0;                                            %  ���������С(��)
        selection.Font.Bold = 0;                                               %  ��������Ӵ�
        selection.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';       %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
        selection.ParagraphFormat.SpaceBefore = 0;                             %  ��ǰ���(��)
        selection.ParagraphFormat.SpaceAfter = 0;                              %  �κ���(��)
        tabstop = selection.ParagraphFormat.TabStops.Add( leftTabstop );       %  ���õڶ����Ʊ��λ��
        tabstop.Alignment = 'wdAlignTabLeft';
        selection.MoveDown;
        selection.TypeParagraph;
    elseif exportFlyPage == 0
        disp('    �������ҳ��ע�����');
    else
        error('�����ҳ��ע����� ѡ�����');
    end
    
    
end

%%%%%%  ���������ҳ�ļ�  %%%%%%
disp('    ���������ҳ�ĵ�');
bgmc = strcat( '������ҳ', bgbh{beginNo}, '��' , bgbh{endNo}  );              %  �����ⱨ������
filespec= fullfile( pwd , '\��ⱨ��\' , bgmc );                              %  ��������ļ����·��
try
    document.SaveAs(filespec);
catch
    document.SaveAs2(filespec);
end

if exportAsPDF == 1
    if  word.Version >= 12
        disp('    ���PDF��ʽ�ļ�');
        document.ExportAsFixedFormat( filespec , 17 );                        %  ���ΪPDF��ʽ   17 - PDF��ʽ��18 - XPS��ʽ
    end
end

document.Close;
% word.Quit;

end


function [] = reportPage( beginNo , endNo , scgsytspec_3 , scgsytspec_4 , exportAsPDF , supervisionPattern , ...     %   ����ҳ�������
    bgxh , bgbh , jgwmc , dtbh , jzbh , concreteStrength , sjzc , sczc , sjzj , gzfl , gzrq , jcrq , jcjg , ...
    qxfx , cggs , cgjj1_2 , cgjj1_3 , cgjj2_3 , cgjj1_4 , cgjj2_4 , cgjj3_4 , fjys , ...
    cdjj , yqsb , jcbz , jzbg , gcmc , wtdw , jcnr , jclb , bxry , jcry , fhry , ...
    shry , pzry , pzjb , jcdw , dwdz , dwdh , yzbm , dwcz , bgrq , bgys , ...
    jsdw , kcdw , sjdw , zjb , jldw , zdjl , sgdw , jcxs , sbjgxs ,  ...
    jcxm , jcff , jcsl , gcgk , dztj , sggc , jcgc , zwpmbzt , bz , ...
    zysx , sjzTbg , sjzBbg , sczTbg , sczBbg )
%  reportPage ��������ҳ���    v20180606


%%  ����word�ĵ��Ļ���������Ϣ

paperSize = 'wdPaperA4';                                                    %  ����ֽ�Ŵ�С    'wdPaperA4'     'wdPaperA3'   'wdPaperB5' ......
orientation = 'wdOrientPortrait';                                           %  ����ֽ�ŷ���   'wdOrientPortrait'  'wdOrientLandscape'
margin = [65.2, 53.85, 53.85, 53.85];                                        %  ����ҳ�߾࣬��λpoint(��)    [TopMargin, BottomMargin, LeftMargin, RightMargin]
headerFooterDistance = [42.55, 49.6];                                       %  ����ҳüҳ�ž�߽磬��λpoint(��)   [ҳü��ҳ��]
columWidth_1 = [127.6, 90.6, 90.6, 127.6];                                  %  ǩ��ҳ��Ԫ����(��)
rowHeight_1 = [41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41];  %  ǩ��ҳ��Ԫ��߶�(��)
height_2 = 28.4;                                                            %  ��1  ���̼���ÿ�и߶�(��)
width_2 = [111.4 , 331.6];                                                  %  ��1  ���̼���ÿ�п��(��)
height_3 = [40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40 ];          %  ��2  ������һ����ÿ�и߶�(��)
width_3 = [100 , 100, 100, 100, 100];                                       %  ��2  ������һ����ÿ�п��(��)
picsize_1 = [487.7 , 304.8 , 1];                                            %  ͼ1  ׮λƽ�沼��ͼͼƬ�ߴ���ֵ��487.7��304.8����  [ ����(��)���Ƿ񱣳�ԭͼƬ��߱�(0/1����) ]
picsize_2 = [ 90 , 70 ];                                                    %  �����λ��ʾ��ͼ ͼƬ�ߴ磬    [ �� �� ��(��) ]



%%  ����Microsoft Word ������

disp('    ���� Microsoft Word ������');

try                                                                         %  ����actxserver��������Microsoft Word ������
    word = actxGetRunningServer('Word.Application');                        %  ��word�������Ѿ��򿪣���������Word
catch
    word = actxserver('Word.Application');                                  %  ��word������û�д򿪣�����һ�������ؾ��Word
end

set(word,'Visible',0);                                                      %  ���÷����������Ϊ�ɼ�״̬, 1 - �ɼ��� 0 - ���ɼ�
document = invoke(word.Documents,'Add');                                    %  ��һ���հ��ĵ�

% if exist(filespec,'file')
%     document = invoke(word.Documents,'Open',filespec);                    %  ���ļ����ڣ��򿪸��ļ��������½�һ���ļ�
% else
%     document = invoke(word.Documents,'Add');                              %  ���ļ������ڣ��½�һ���ļ�
%     document.SaveAs2(filespec);                                           %  �����ĵ�
% end

selection = word.Selection;                                                 %  ����Word.Selection�ӿھ��
content = document.Content;                                                 %  ����Word.Documents.Content�ӿھ��


%%  ҳ������

disp('    �ĵ�ҳ������');

document.PageSetup.PaperSize = paperSize;                                   % ����ֽ�Ŵ�С��A4��A3��B5
document.PageSetup.LayoutMode = 'wdLayoutModeDefault';                      % �����ĵ�����Ϊ������
document.PageSetup.Orientation = orientation;                               % ����ֽ�ŷ���
document.PageSetup.TopMargin = margin(1);                                   % ҳ��߾�����
document.PageSetup.BottomMargin = margin(2);
document.PageSetup.LeftMargin = margin(3);
document.PageSetup.RightMargin = margin(4);
document.PageSetup.HeaderDistance = headerFooterDistance(1);                %  ����ҳü��߽����
document.PageSetup.FooterDistance = headerFooterDistance(2);                %  ����ҳ�ž�߽����
headerFooterTabstop = document.PageSetup.PageWidth - margin(3) - margin(4); %  ����ҳü�Ҷ˶����Ʊ��λ��


%% ��ʽ����

disp('    �ĵ���ʽ����');

try
    userStyles = document.Styles.Item('����');
catch
    userStyles = document.Styles.Add('����');
end                                                                        %  ���á����ġ���ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 12.0;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';           %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
% userStyles.ParagraphFormat.LeftIndent = 0;                                %  �������������(��)
% userStyles.ParagraphFormat.FirstlineIndent = 24;                          %  ��������(��)  ��ֵ��������
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 2;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';             %  �����о�    'wdLineSpaceSingle' 'wdLineSpace1pt5'



try
    userStyles = document.Styles.Add('������');
catch
    userStyles = document.Styles.Item('������');
end                                                                         %  ���á������⡱��ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 16.0;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  ���ö��뷽ʽ'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 8.0;                               %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 8.0;                                %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';             %  �����о�    'wdLineSpaceSingle' 'wdLineSpace1p


try
    userStyles = document.Styles.Add('���� 1');
catch
    userStyles = document.Styles.Item('���� 1');
end                                                                         %  ���á����� 2����ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 15.0;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';              %  ���ö��뷽ʽ'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevel1';                %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 7.5;                               %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';             %  �����о�    'wdLineSpaceSingle' 'wdLineSpace1pt5'


try
    userStyles = document.Styles.Add('���� 2');
catch
    userStyles = document.Styles.Item('���� 2');
end                                                                         %  ���á����� 2����ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 14.0;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';              %  ���ö��뷽ʽ'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel ='wdOutlineLevel2';                 %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 1;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';


try
    userStyles = document.Styles.Add('��ע');
catch
    userStyles = document.Styles.Item('��ע');
end                                                                         %  ���á���ע����ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 10.5;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  ���ö��뷽ʽ'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';


try
    userStyles = document.Styles.Add('���');
catch
    userStyles = document.Styles.Item('���');
end                                                                         %  ���á������ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 12.0;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  ���ö��뷽ʽ'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  �����о�    'wdLineSpaceSingle' 'wdLineSpace1pt5'


try
    userStyles = document.Styles.Add('ǩ��ҳ');
catch
    userStyles = document.Styles.Item('ǩ��ҳ');
end                                                                         %  ���á�ǩ��ҳ����ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 14.0;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  ���ö��뷽ʽ'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  �����о�    'wdLineSpaceSingle' 'wdLineSpace1pt5'


try
    userStyles = document.Styles.Item('ҳü');
catch
    userStyles = document.Styles.Add('ҳü');
end                                                                         %  ���á�ҳü����ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 10.5;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';              %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)  ��ֵ��������
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  �����о�    'wdLineSpaceSingle' 'wdLineSpace1pt5'
userStyles.ParagraphFormat.TabStops.ClearAll;


try
    userStyles = document.Styles.Item('ҳ��');
catch
    userStyles = document.Styles.Add('ҳ��');
end                                                                         %  ���á�ҳ�š���ʽ
userStyles.Font.Name = '����';                                              %  ����Ĭ������Ϊ����
userStyles.Font.NameFarEast = '����';                                       %  ������������Ϊ����
userStyles.Font.NameAscii = 'Times New Roman';                              %  ����Ascii����ΪTimes New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  ���������ַ�����ΪTimes New Roman
userStyles.Font.Size = 10.5;                                                %  ���������С(��)
userStyles.Font.Bold = 0;                                                   %  ��������Ӵ�
userStyles.Font.Italic = 0;                                                 %  ��������б��
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  ���ö��뷽ʽ 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  ���ô�ټ���'wdOutlineLevelBodyText'��'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  �������������(��)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  ��������(��)  ��ֵ��������
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  �������������(�ַ�)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  ��������(�ַ�)����Ϊ������������Ϊ��������
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  ��ǰ���(��)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  �κ���(��)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  �����о�    'wdLineSpaceSingle' 'wdLineSpace1pt5'



%%  �������

numTables = 0;                                                              %  ������
selection.Style = '����';
for ii = beginNo : 1 : endNo
    
    disp(['    ������ ', bgxh{ii} ,' ������( ' , bgbh{ii},' )����ҳ']);
    if  bz{ii} ~= char(32)                                                      %  �����ע�ǿո���Ļ���������ע������
        disp( [ '              ��ע��'  , bz{ii} ] );
    end
    
    %%%%%%%%%%%%%%%   ǩ��ҳ ���   %%%%%%%%%%%%%%%%%%
    
    
    if ii ~= beginNo
        selection.InsertBreak(2);                                           %  �����ҳ��0��ֽڷ�2
    end
    
    selection.TypeParagraph;                                                %  �س�������һ��
    selection.Start = content.end;                                          %  ��ѡ���������ʼλ�ö�λ������ĩβ
    document.Tables.Add(selection.Range,16,4);                              %  �����һҳ���
    numTables = numTables + 1;
    DTI = document.Tables.Item( numTables );                                %  ��ȡ�½������
    
    DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';                     %  ������߿����ͣ���ѯ��DTI.Borders.set('OutsideLineStyle')
    DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';                      %  ������߿��߿�
    DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';
    DTI.Borders.InsideLineWidth = 'wdLineWidth050pt';
    DTI.Rows.Alignment = 'wdAlignRowCenter';                                %  ������������ˮƽ���뷽ʽ����ѯ��set����
    
    
    for mm = 1:1:16
        DTI.Rows.Item(mm).Height = rowHeight_1(mm);                         %  ����ǩ��ҳ��Ԫ��߶�
        for nn = 1:1:4
            DTI.Columns.Item(nn).Width = columWidth_1(nn);                  %  ����ǩ��ҳ��Ԫ����
            DTI.Cell(mm,nn).Range.Paragraphs.Alignment = 'wdAlignParagraphCenter';  %  ���õ�Ԫ��ˮƽ���ж���
            DTI.Cell(mm,nn).VerticalAlignment = 'wdCellAlignVerticalCenter';   %  ���õ�Ԫ����뷽ʽ
            DTI.Cell(mm,nn).Range.ParagraphFormat.Style = 'ǩ��ҳ';          %   ������ʽΪ��ǩ��ҳ��
            
        end
    end
    
    
    DTI.Cell(1,2).Merge(DTI.Cell(1,4));                                     %  �ϲ���Ԫ�񣬵�1��2��4�ϲ�
    DTI.Cell(2,2).Merge(DTI.Cell(2,4));
    DTI.Cell(3,2).Merge(DTI.Cell(3,4));
    DTI.Cell(4,2).Merge(DTI.Cell(4,4));
    DTI.Cell(5,2).Merge(DTI.Cell(5,4));
    DTI.Cell(6,2).Merge(DTI.Cell(6,4));
    DTI.Cell(7,2).Merge(DTI.Cell(7,4));
    DTI.Cell(8,2).Merge(DTI.Cell(8,4));
    DTI.Cell(9,2).Merge(DTI.Cell(9,4));
    DTI.Cell(10,2).Merge(DTI.Cell(10,4));
    DTI.Cell(11,2).Merge(DTI.Cell(11,4));
    DTI.Cell(12,2).Merge(DTI.Cell(12,4));
    DTI.Cell(13,2).Merge(DTI.Cell(13,4));
    DTI.Cell(15,1).Merge(DTI.Cell(15,4));
    DTI.Cell(16,1).Merge(DTI.Cell(16,4));
    
    
    DTI.Cell(1,1).Range.Text = '�� �� �� ��';                                %  ¼��������
    DTI.Cell(2,1).Range.Text = '�� �� �� ��';
    DTI.Cell(3,1).Range.Text = 'ί �� �� λ';
    DTI.Cell(4,1).Range.Text = '�� �� �� ��';
    DTI.Cell(5,1).Range.Text = '�� �� �� ��';
    DTI.Cell(6,1).Range.Text = '�� �� �� Ա';
    DTI.Cell(7,1).Range.Text = '�� д �� Ա';    
    DTI.Cell(8,1).Range.Text = '�� �� (ǩ��)';
    DTI.Cell(9,1).Range.Text = '�� �� (ǩ��)';
    DTI.Cell(10,1).Range.Text = '�� ׼ (ǩ��)';
    DTI.Cell(11,1).Range.Text = '�� �� �� λ';
    DTI.Cell(12,1).Range.Text = '��      ַ';
    DTI.Cell(13,1).Range.Text = '��      ��';
    DTI.Cell(14,1).Range.Text = '�� �� �� ��';
    DTI.Cell(14,3).Range.Text = '��  ��';
    
    for mm = 1:14                                                           %  ����Ӻ�
        DTI.Cell(mm,1).Range.Font.Bold = 1;
    end
    DTI.Cell(14,3).Range.Font.Bold = 1;
    
    DTI.Cell(1,2).Range.Text = bgbh{ii};      %  ������
    DTI.Cell(2,2).Range.Text = gcmc{ii};      %  ��������
    DTI.Cell(3,2).Range.Text = wtdw{ii};      %  ί�е�λ
    DTI.Cell(4,2).Range.Text = jcnr{ii};      %  �������
    DTI.Cell(5,2).Range.Text = jclb{ii};      %  ������
    DTI.Cell(6,2).Range.Text = jcry{ii};      %  �����Ա
    DTI.Cell(7,2).Range.Text = bxry{ii};      %  ��д��Ա    
    DTI.Cell(8,2).Range.Text = fhry{ii};      %  ������Ա
    DTI.Cell(9,2).Range.Text = shry{ii};      %  �����Ա
    DTI.Cell(10,2).Range.Text = pzry{ii};     %  ��׼��Ա
    DTI.Cell(11,2).Range.Text = jcdw{ii};     %  ��ⵥλ
    DTI.Cell(12,2).Range.Text = dwdz{ii};     %  ��λ��ַ
    DTI.Cell(13,2).Range.Text = dwdh{ii};     %  ��λ�绰
    DTI.Cell(14,2).Range.Text = yzbm{ii};     %  �������
    DTI.Cell(14,4).Range.Text = dwcz{ii};     %  ��λ����
    
    selection.Tables.Item( 1 ).Cell(15,1).Select();
    selection.Text = '�������ڣ�';
    selection.Font.Bold = 1;
    selection.EndKey();
    selection.Text =  datestr( datenum( strtok(bgrq{ii}) ), 'yyyy��mm��dd��' );
    selection.Font.Bold = 0;
    
    selection.Tables.Item( 1 ).Cell(16,1).Select();
    selection.Text = '����ҳ��';
    selection.Font.Bold = 1;
    selection.EndKey();
    selection.Text = strcat('��1ҳ ��', bgys{ii}, 'ҳ');
    selection.Font.Bold = 0;
    
    selection.Start = content.end;                                          %  ��ѡ���������ʼλ�ö�λ������ĩβ
    selection.InsertBreak(2);                                               %  �����ҳ��0��ֽڷ�2
    
    
    %%%%%%%%%%%%%%%%%   ǩ��ҳ ������   %%%%%%%%%%%%%%%%%%
    
    
    
    
    %%%%%%%%%%%%%%%%%        ����ҳ      %%%%%%%%%%%%%%%%%%%%%
    
    selection.Start = content.end;                                         %  ��ѡ���������ʼλ�ö�λ������ĩβ
    selection.Text = strcat(jcnr{ii} , '����')';                           %  ����������  %  selection.Text = '��������ע׮׮�������Գ�����͸�䷨��ⱨ��';
    selection.Style = '������';                                            %   ��ʽ'������'
    selection.MoveDown;                                                    %  ������Ƶ���ѡ��������,�����־�ĺ��棬Ҳ����һ�ε���ǰ��;������EndKey�����������棬�����־��ǰ��
    
    
    selection.TypeParagraph;                                               %  �س�������һ��
    selection.Text = '1���̼��';
    selection.Style = '���� 1';
    selection.MoveDown;                                                    %  ������Ƶ���ѡ��������
    
    
    selection.TypeParagraph;                                               %  �س�������һ��
    selection.Text = '�� 1  ���̼���';
    selection.Style = '��ע';
    selection.MoveDown;                                                    %  ������Ƶ���ѡ��������
    
    selection.TypeParagraph;
    if supervisionPattern == 1                                             %  ���ݼ���ģʽѡ�񹤳̼�����ʽ�� 1 - �ܼ�졢פ�ؼ���ģʽ�� 2 - ��һ����λģʽ
        document.Tables.Add(selection.Range,12,2);                              %  �����һҳ���
        numTables = numTables + 1;
        DTI = document.Tables.Item( numTables );
        
        for mm = 1:1:12
            DTI.Rows.Item(mm).Height = height_2;                                %  ���ñ�1��Ԫ��߶�
            for nn = 1:1:2
                DTI.Columns.Item(nn).Width = width_2(nn);                       %  ���ñ�1��Ԫ����
                DTI.Cell(mm,nn).Range.Paragraphs.Alignment = 'wdAlignParagraphCenter';  %  ���õ�Ԫ��ˮƽ���ж���
                DTI.Cell(mm,nn).VerticalAlignment = 'wdCellAlignVerticalCenter';        %  ���õ�Ԫ����뷽ʽ
                DTI.Cell(mm,nn).Range.ParagraphFormat.Style = '���';            %   ������ʽΪ�����
            end
        end
        
        DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';                     %  ������߿����ͣ���ѯ��DTI.Borders.set('OutsideLineStyle')
        DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';                      %  ������߿��߿�
        DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';
        DTI.Borders.InsideLineWidth = 'wdLineWidth050pt';
        DTI.Rows.Alignment = 'wdAlignRowCenter';                                %  ���õ�Ԫ��ˮƽ���ж���
        DTI.Cell(1,1).Range.Text = '��������';                                   %  ¼��������
        DTI.Cell(2,1).Range.Text = '���赥λ';
        DTI.Cell(3,1).Range.Text = '���쵥λ';
        DTI.Cell(4,1).Range.Text = '��Ƶ�λ';
        DTI.Cell(5,1).Range.Text = '�ܼ��';
        DTI.Cell(6,1).Range.Text = 'פ�ؼ���';
        DTI.Cell(7,1).Range.Text = 'ʩ����λ';
        DTI.Cell(8,1).Range.Text = '������ʽ';
        DTI.Cell(9,1).Range.Text = '�ϲ��ṹ��ʽ';
        DTI.Cell(10,1).Range.Text = '�����Ŀ';
        DTI.Cell(11,1).Range.Text = '��ⷽ��';
        DTI.Cell(12,1).Range.Text = '�������';
        DTI.Cell(1,2).Range.Text = gcmc{ii};
        DTI.Cell(2,2).Range.Text = jsdw{ii};
        DTI.Cell(3,2).Range.Text = kcdw{ii};
        DTI.Cell(4,2).Range.Text = sjdw{ii};
        DTI.Cell(5,2).Range.Text = zjb{ii};
        DTI.Cell(6,2).Range.Text = zdjl{ii};
        DTI.Cell(7,2).Range.Text = sgdw{ii};
        DTI.Cell(8,2).Range.Text = jcxs{ii};
        DTI.Cell(9,2).Range.Text = sbjgxs{ii};
        DTI.Cell(10,2).Range.Text = jcxm{ii};
        DTI.Cell(11,2).Range.Text = jcff{ii};
        DTI.Cell(12,2).Range.Text = jcsl{ii};
        selection.Start = content.end;
        
        
    elseif supervisionPattern == 2                                         %  ���ݼ���ģʽѡ�񹤳̼�����ʽ�� 1 - �ܼ�졢פ�ؼ���ģʽ�� 2 - ��һ����λģʽ
        document.Tables.Add(selection.Range,11,2);                              %  �����һҳ���
        numTables = numTables + 1;
        DTI = document.Tables.Item( numTables );
        
        for mm = 1:1:11
            DTI.Rows.Item(mm).Height = height_2;                                %  ���ñ�1��Ԫ��߶�
            for nn = 1:1:2
                DTI.Columns.Item(nn).Width = width_2(nn);                       %  ���ñ�1��Ԫ����
                DTI.Cell(mm,nn).Range.Paragraphs.Alignment = 'wdAlignParagraphCenter';  %  ���õ�Ԫ��ˮƽ���ж���
                DTI.Cell(mm,nn).VerticalAlignment = 'wdCellAlignVerticalCenter';        %  ���õ�Ԫ����뷽ʽ
                DTI.Cell(mm,nn).Range.ParagraphFormat.Style = '���';            %   ������ʽΪ�����
            end
        end
        
        DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';                     %  ������߿����ͣ���ѯ��DTI.Borders.set('OutsideLineStyle')
        DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';                      %  ������߿��߿�
        DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';
        DTI.Borders.InsideLineWidth = 'wdLineWidth050pt';
        DTI.Rows.Alignment = 'wdAlignRowCenter';                                %  ���õ�Ԫ��ˮƽ���ж���
        DTI.Cell(1,1).Range.Text = '��������';                                   %  ¼��������
        DTI.Cell(2,1).Range.Text = '���赥λ';
        DTI.Cell(3,1).Range.Text = '���쵥λ';
        DTI.Cell(4,1).Range.Text = '��Ƶ�λ';
        DTI.Cell(5,1).Range.Text = '����λ';
        DTI.Cell(6,1).Range.Text = 'ʩ����λ';
        DTI.Cell(7,1).Range.Text = '������ʽ';
        DTI.Cell(8,1).Range.Text = '�ϲ��ṹ��ʽ';
        DTI.Cell(9,1).Range.Text = '�����Ŀ';
        DTI.Cell(10,1).Range.Text = '��ⷽ��';
        DTI.Cell(11,1).Range.Text = '�������';
        DTI.Cell(1,2).Range.Text = gcmc{ii};
        DTI.Cell(2,2).Range.Text = jsdw{ii};
        DTI.Cell(3,2).Range.Text = kcdw{ii};
        DTI.Cell(4,2).Range.Text = sjdw{ii};
        DTI.Cell(5,2).Range.Text = jldw{ii};
        DTI.Cell(6,2).Range.Text = sgdw{ii};
        DTI.Cell(7,2).Range.Text = jcxs{ii};
        DTI.Cell(8,2).Range.Text = sbjgxs{ii};
        DTI.Cell(9,2).Range.Text = jcxm{ii};
        DTI.Cell(10,2).Range.Text = jcff{ii};
        DTI.Cell(11,2).Range.Text = jcsl{ii};
        selection.Start = content.end;
        
    end
    
    
    
    selection.Text = '2���̸ſ���������������';
    selection.Style = '���� 1';
    selection.MoveDown;
    
    
    selection.TypeParagraph;                                                %  �س�������һ��
    selection.Text = '2.1���̸ſ�';
    selection.Style = '���� 2';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  �س�������һ��
    gcgk1 = strrep(gcgk{ii}, 'wtdw', wtdw{ii});                             %  ������'wtdw'�滻Ϊί�е�λwtdw{ii}
    gcgk2 = strrep(gcgk1, 'jcdw' , jcdw{ii} );                              %  ������'jcdw'�滻Ϊ��ⵥλjcdw{ii}
    gcgk3 = strrep(gcgk2, 'gcmc' , erase(gcmc{ii}, newline));               %  ɾ�����������ַ����еĻ��з���Ȼ������'gcmc'�滻Ϊ��������gcmc{ii}
                                                                            %   ��MatLab�汾���ã���gcgk3 = strrep(gcgk2, 'gcmc' , strrep(gcmc{ii},char(10),'')); ��
    gcgk4 = strrep(gcgk3, 'jgwmc' , jgwmc{ii} );                            %  ������'jgwmc'�滻Ϊ�ṹ������jgcmc{ii}
    selection.Text = gcgk4;
    selection.Style = '����';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  �س�������һ��
    pilepositionfilespec = fullfile(pwd,'\�����������ļ�\',zwpmbzt{ii});     %   ����׮λƽ�沼��ͼ��·��������
    handle1 = selection.InlineShape.AddPicture(pilepositionfilespec);
    if picsize_1(3) == 1                                                    %  ����ͼƬ��߱�����ͼƬ
        scaling = min( picsize_1(1)/handle1.Width , picsize_1(2)/handle1.Height );  %  ͼ1�����ű���
        handle1.Width = handle1.Width * scaling;                            % �ȱ�������ͼ1�Ŀ���
        handle1.Height = handle1.Height * scaling;
    elseif picsize_1(3) == 0                                                %   ������ͼƬ��߱�����ͼƬ
        handle1.Width = picsize_1(1);
        handle1.Height = picsize_1(2);
    else
        error('picsize_1ͼƬ������߱Ȳ�����������(0��1)�����޸ĺ����ԣ�');
    end
    selection.Style = '���';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  �س�������һ��
    selection.Text = 'ͼ 1  ׮λƽ�沼��ͼ';
    selection.Style = '��ע';
    selection.MoveDown;                                                    %  ������Ƶ���ѡ��������
    
    selection.TypeParagraph;                                                %  �س�������һ��
    selection.Text = '2.2������������';
    selection.Style = '���� 2';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  �س�������һ��
    selection.Text = dztj{ii};
    selection.Style = '����';
    selection.MoveDown;
    
    selection.Start = content.end;                                          %  ��ѡ���������ʼλ�ö�λ������ĩβ
    selection.InsertBreak(0);                                               %  �����ҳ��0��ֽڷ�2
    
    
    selection.Text = '3������һ����';
    selection.Style = '���� 1';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  �س�������һ��
    selection.Text = '�� 2  ������һ����';
    selection.Style = '��ע';
    selection.MoveDown;
    
    
    selection.TypeParagraph;
    document.Tables.Add(selection.Range,13,5);                              %  ������2
    numTables = numTables + 1;
    DTI = document.Tables.Item( numTables );
    
    for mm = 1:1:13
        DTI.Rows.Item(mm).Height = height_3(mm);                             %  ���ñ�2��Ԫ��߶�
        for nn = 1:1:5
            DTI.Columns.Item(nn).Width = width_3(nn);                        %  ���ñ�2��Ԫ����
            DTI.Cell(mm,nn).Range.Paragraphs.Alignment = 'wdAlignParagraphCenter';  %  ���õ�Ԫ��ˮƽ���ж���
            DTI.Cell(mm,nn).VerticalAlignment = 'wdCellAlignVerticalCenter';        %  ���õ�Ԫ����뷽ʽ
            DTI.Cell(mm,nn).Range.ParagraphFormat.Style = '���';            %   ������ʽΪ�����
        end
    end
    for mm = 1:2:7                                                         %  ����Ӻ�
        DTI.Cell(mm,1).Range.Font.Bold = 1;
        DTI.Cell(mm,2).Range.Font.Bold = 1;
        DTI.Cell(mm,3).Range.Font.Bold = 1;
        DTI.Cell(mm,4).Range.Font.Bold = 1;
        DTI.Cell(mm,5).Range.Font.Bold = 1;
    end
    DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';                     %  ������߿����ͣ���ѯ��DTI.Borders.set('OutsideLineStyle')
    DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';                      %  ������߿��߿�
    DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';
    DTI.Borders.InsideLineWidth = 'wdLineWidth050pt';
    DTI.Rows.Alignment = 'wdAlignRowCenter';                                %  ��
    DTI.Cell(8,1).Merge(DTI.Cell(13,1));                                    %  �ϲ���Ԫ�񣬵�1��2��4�ϲ�
    DTI.Cell(8,2).Merge(DTI.Cell(13,2));
    DTI.Cell(8,3).Merge(DTI.Cell(13,3));
    DTI.Cell(7,4).Merge(DTI.Cell(7,5));
    DTI.Cell(1,1).Range.Text = '��������';                                  %  ¼��������
    DTI.Cell(2,1).Range.Text = gcmc{ii};
    DTI.Cell(1,2).Range.Text = '�ṹ������';
    DTI.Cell(2,2).Range.Text = jgwmc{ii};
    DTI.Cell(1,3).Range.Text = '�����豸';
    DTI.Cell(2,3).Range.Text = yqsb{ii};
    DTI.Cell(1,4).Range.Text = '����׼';
    DTI.Cell(2,4).Range.Text = jcbz{ii};
    DTI.Cell(1,5).Range.Text = '��׮���';
    DTI.Cell(2,5).Range.Text = jzbh{ii};    
    DTI.Cell(3,1).Range.Text = '���׮�����(m)';                                  %  ¼��������
    DTI.Cell(4,1).Range.Text = sjzTbg{ii};
    DTI.Cell(3,2).Range.Text = '���׮�˱��(m)';
    DTI.Cell(4,2).Range.Text = sjzBbg{ii};
    DTI.Cell(3,3).Range.Text = '���׮��(m)';
    DTI.Cell(4,3).Range.Text = sjzc{ii};
    DTI.Cell(3,4).Range.Text = '���׮��(mm)';                                %  ¼��������
    DTI.Cell(4,4).Range.Text = sjzj{ii};
    DTI.Cell(3,5).Range.Text = '������ǿ�ȵȼ�';
    DTI.Cell(4,5).Range.Text = concreteStrength{ii};
    DTI.Cell(5,1).Range.Text = '��ע����';
    DTI.Cell(6,1).Range.Text = datestr( datenum( strtok(gzrq{ii}) ), 'yyyy-mm-dd' );
    DTI.Cell(5,2).Range.Text = '�������';
    DTI.Cell(6,2).Range.Text = datestr( datenum( strtok(jcrq{ii}) ), 'yyyy-mm-dd' );
    DTI.Cell(5,3).Range.Text = 'ʵ��׮��(m)';
    DTI.Cell(6,3).Range.Text = sczc{ii};
    DTI.Cell(5,4).Range.Text = 'ʵ��׮�����(m)';
    DTI.Cell(6,4).Range.Text = sczTbg{ii};
    
    DTI.Cell(5,5).Select();
    selection.Text = '��ע����������(m';
    selection.EndKey;
    selection.Text = '3';
    selection.Font.Superscript = 1;
    selection.EndKey;
    selection.Text = ')';
    selection.Font.Superscript = 0;
    
    DTI.Cell(6,5).Range.Text = gzfl{ii};
    DTI.Cell(7,1).Range.Text = '����ܷ�λʾ��ͼ';
    DTI.Cell(7,2).Range.Text = '�����';
    DTI.Cell(8,2).Range.Text = jcjg{ii}';
    DTI.Cell(7,3).Range.Text = 'ȱ���������';
    DTI.Cell(8,3).Range.Text = qxfx{ii}';
    DTI.Cell(7,4).Range.Text = '��ܼ��(mm)';
    DTI.Cell(8,4).Range.Text = '1-2';
    DTI.Cell(8,5).Range.Text = cgjj1_2{ii};
    DTI.Cell(9,4).Range.Text = '1-3';
    DTI.Cell(9,5).Range.Text = cgjj1_3{ii};
    DTI.Cell(10,4).Range.Text = '2-3';
    DTI.Cell(10,5).Range.Text = cgjj2_3{ii};
    
    if str2double(cggs{ii}) == 3                                         %  ������ ����ܷ�λʾ��ͼ
        selection.Tables.Item( 1 ).Cell(8,1).Select();
        selection.EndKey();
        handle1 = selection.InlineShapes.AddPicture(scgsytspec_3);
        handle1.Width = picsize_2(1);
        handle1.Height = picsize_2(2);
        DTI.Cell(11,4).Range.Text = '/';
        DTI.Cell(11,5).Range.Text = '/';
        DTI.Cell(12,4).Range.Text = '/';
        DTI.Cell(12,5).Range.Text = '/';
        DTI.Cell(13,4).Range.Text = '/';
        DTI.Cell(13,5).Range.Text = '/';
    elseif str2double(cggs{ii}) == 4                                     %  �ĸ��� ����ܷ�λʾ��ͼ
        selection.Tables.Item( 1 ).Cell(8,1).Select();
        selection.EndKey();
        handle1 = selection.InlineShapes.AddPicture(scgsytspec_4);
        handle1.Width = picsize_2(1);
        handle1.Height = picsize_2(2);
        DTI.Cell(11,4).Range.Text = '1-4';
        DTI.Cell(11,5).Range.Text = cgjj1_4{ii};
        DTI.Cell(12,4).Range.Text = '2-4';
        DTI.Cell(12,5).Range.Text = cgjj2_4{ii};
        DTI.Cell(13,4).Range.Text = '3-4';
        DTI.Cell(13,5).Range.Text = cgjj3_4{ii};
    end
    selection.Start = content.end;
    
    selection.Text = strcat('������ʵ�����ߡ�ͼ����', fjys{ii}, 'ҳ��');
    selection.Style = '����';
    selection.MoveDown;
    
    
end



%%%%%%%%%    ����ҳü    %%%%%%%%%%
for mm = 1 : 2 : document.Sections.Count                                    %  �м���2Ϊÿ2��ִ��һ��ѭ��
    disp(['    ������� ' , bgxh{beginNo+(mm-1)/2} , ' ������( ', bgbh{beginNo+(mm-1)/2} ,' )ҳü']);
    document.Sections.Item(mm).Headers.Item('wdHeaderFooterPrimary').LinkToPrevious = 0;                                                                 %  ���ñ��ڲ����ӵ�ǰһ��ҳü
    document.Sections.Item(mm).Headers.Item('wdHeaderFooterPrimary').Range.Text='';                                                                      %  ���ñ���ҳü����Ϊ''
    document.Sections.Item(mm).Headers.Item('wdHeaderFooterPrimary').Range.ParagraphFormat.Borders.Item('wdBorderBottom').LineStyle = 'wdLineStyleNone'; %  ���ñ��ڶ����»���Ϊ��
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').LinkToPrevious = 0;                                                               %  ���õڶ��ڲ����ӵ�ǰһ��ҳü
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.Text = strcat(jcnr{beginNo+(mm-1)/2},'����',09,bgbh{beginNo+(mm-1)/2});   %   strcat('��������׮�����Գ��������',09,bgbh{beginNo+(mm-1)/2});         %   ���õڶ���ҳü����
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.Style = 'ҳü';                                                            %   ���õڶ���ҳü��ʽΪ'ҳü1'
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.ParagraphFormat.Borders.Item('wdBorderBottom').LineStyle = 'wdLineStyleThinThickSmallGap';   %  ���õڶ��ڶ����»���
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.ParagraphFormat.Borders.Item('wdBorderBottom').LineWidth = 'wdLineWidth150pt';               %  ���õڶ��ڶ����»��߿��
    tabstop = document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.ParagraphFormat.TabStops.Add( headerFooterTabstop );                               %  ���õڶ����Ʊ��λ��
    tabstop.Alignment = 'wdAlignTabRight';                                                                                                                                %  �����Ʊ���Ķ��뷽ʽ
end

%%%%%%%%    ����ҳ��    %%%%%%%%%%
for nn = 1 : 2 : document.Sections.Count
    disp(['    ������� ' , bgxh{beginNo+(nn-1)/2} , ' ������( ', bgbh{beginNo+(nn-1)/2} ,' )ҳ��']);
    document.Sections.Item(nn).Footers.Item('wdHeaderFooterPrimary').LinkToPrevious = 0;            %  ���ñ���ҳ�Ų�������һ��ҳ��
    document.Sections.Item(nn).Footers.Item('wdHeaderFooterPrimary').Range.Text='';                 %  ���ñ���ҳ������
    document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').LinkToPrevious = 0;          %  ���õڶ���ҳ�Ų�������һ��ҳ��
    document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').Range.Fields.Add(document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').Range,[],'Page');   %  ����ڶ���ҳ��ҳ��
    document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').Range.Select;
    selection.Range.InsertBefore( '��');
    selection.Range.InsertAfter(strcat( 'ҳ ��',bgys{beginNo+(nn-1)/2}, 'ҳ'));
    document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').Range.Style = 'ҳ��';
    
    document.Sections.Item(nn).Footers.Item('wdHeaderFooterPrimary').PageNumbers.RestartNumberingAtSection = 1;  %  ���ñ���ҳ���ؿ�ʼ����
    document.Sections.Item(nn).Footers.Item('wdHeaderFooterPrimary').PageNumbers.StartingNumber = 1;             %  ���ñ���ҳ���ؿ�ʼ����Ϊ1
end
word.ActiveWindow.View.Type='wdPrintView';
document.ActiveWindow.ActivePane.View.SeekView ='wdSeekMainDocument';
word.ActiveWindow.View.Type='wdPrintView';


%%%%%%  �����ļ�  %%%%%%
disp('    ���汨��ҳ�ĵ�');
bgmc = strcat( '����ҳ', bgbh{beginNo}, '��' , bgbh{endNo}  );                         %  �����ⱨ������
filespec= fullfile( pwd , '\��ⱨ��\' , bgmc );                              %  ��������ļ����·��
try
    document.SaveAs(filespec);
catch
    document.SaveAs2(filespec);
end

if exportAsPDF == 1
    if  word.Version >= 12
        disp('    ���PDF��ʽ�ļ�');
        document.ExportAsFixedFormat( filespec , 17 );                            %  ���ΪPDF��ʽ   17 - PDF��ʽ��18 - XPS��ʽ
    end
end

document.Close;
% word.Quit;


end


function [ bgxh , bgbh , jgwmc , dtbh , jzbh , concreteStrength , sjzc , sczc , sjzj , gzfl , gzrq , jcrq , jcjg , ...
    qxfx , cggs , cgjj1_2 , cgjj1_3 , cgjj2_3 , cgjj1_4 , cgjj2_4 , cgjj3_4 , fjys , ...
    cdjj , yqsb , jcbz , jzbg , gcmc , wtdw, jcnr, jclb , bxry , jcry , fhry , ...
    shry , pzry , pzjb , jcdw , dwdz , dwdh , yzbm , dwcz , bgrq , bgys , ...
    jsdw , kcdw , sjdw , zjb , jldw , zdjl , sgdw , jcxs , sbjgxs ,  ...
    jcxm , jcff , jcsl , gcgk , dztj , sggc , jcgc , zwpmbzt , bz , ...
    zysx , sjzTbg , sjzBbg , sczTbg , sczBbg ] ...
    = readParameters( raw1A , raw2A )
%%   readParameters   ��ȡexcel�ļ��в���

raw1 = {};
raw2 = {};

for ii = 1 : 1 : size( raw1A , 1 )                                          %  ȥ��Excel�пհ���
    if ~isnan( raw1A{ii,1})                                                 %  �жϵ�ii�е�1�����ַǿ�
        raw1 = [ raw1 ; raw1A(ii,:) ];
    end
end

for ii = 1 : 1 : size( raw2A , 1 )                                          %  ȥ��Excel�пհ���
    if ~isnan( raw2A{ii,1})                                                 %  �жϵ�ii�е�1�����ַǿ�
        raw2 = [ raw2 ; raw2A(ii,:) ];
    end
end

[raw1_m, raw1_n] = size(raw1);
[raw2_m, raw2_n] = size(raw2);

for ii = 1:1:raw1_m
    for jj =1:1:raw1_n
        if isnan(raw1{ ii , jj } )
            raw1{ ii , jj } = char(32);                                    %  ����ļ���û���������Ϣ���Զ�����Ϊ�ո�
        end
    end
end
for ii = 1:1:raw2_m
    for jj =1:1:raw2_n
        if isnan(raw2{ ii , jj } )
            raw2{ ii , jj } = char(32);                                    %  ����ļ���û���������Ϣ���Զ�����Ϊ�ո�
        end
    end
end


bgxh = cell(raw2_m - 1, 1);      %  �������
bgbh = cell(raw2_m - 1, 1);      %  ������
jgwmc = cell(raw2_m - 1, 1);     %  �ṹ������(����)
dtbh = cell(raw2_m - 1, 1);      %  ��̨���
jzbh = cell(raw2_m - 1, 1);      %  ��׮���(׮��)
concreteStrength = cell(raw2_m - 1, 1);      %  ������ǿ�ȵȼ�
sjzc = cell(raw2_m - 1, 1);      %  ���׮��(m)
sczc = cell(raw2_m - 1, 1);      %  ʵ��׮��(m)
sjzj = cell(raw2_m - 1, 1);      %  ���׮��(mm)
sjzTbg = cell(raw2_m - 1, 1);    %  ���׮�����(m)
sjzBbg = cell(raw2_m - 1, 1);    %  ���׮�˱��(m)
sczTbg = cell(raw2_m - 1, 1);    %  ʵ��׮�����(m)
sczBbg = cell(raw2_m - 1, 1);    %  ʵ��׮�˱��(m)
gzfl = cell(raw2_m - 1, 1);      %  ��ע����(m3)
gzrq = cell(raw2_m - 1, 1);      %  ��ע����
jcrq = cell(raw2_m - 1, 1);      %  �������
jcjg = cell(raw2_m - 1, 1);      %  �����
qxfx = cell(raw2_m - 1, 1);      %  ȱ���������
cggs = cell(raw2_m - 1, 1);      %  ��ܸ���
cgjj1_2 = cell(raw2_m - 1, 1);   %  ��ܼ��1-2(mm)
cgjj1_3 = cell(raw2_m - 1, 1);   %  ��ܼ��1-3(mm)
cgjj2_3 = cell(raw2_m - 1, 1);   %  ��ܼ��2-3(mm)
cgjj1_4 = cell(raw2_m - 1, 1);   %  ��ܼ��1-4(mm)
cgjj2_4 = cell(raw2_m - 1, 1);   %  ��ܼ��2-4(mm)
cgjj3_4 = cell(raw2_m - 1, 1);   %  ��ܼ��3-4(mm)
fjys = cell(raw2_m - 1, 1);      %  ����ҳ��
cdjj = cell(raw2_m - 1, 1);      %  �����(m)
yqsb = cell(raw2_m - 1, 1);      %  �����豸
jcbz = cell(raw2_m - 1, 1);      %  ����׼
jzbg = cell(raw2_m - 1, 1);      %  ��׼���
gcmc = cell(raw2_m - 1, 1);      %  ��������
wtdw = cell(raw2_m - 1, 1);      %  ί�е�λ
jcnr = cell(raw2_m - 1, 1);      %  �������
jclb = cell(raw2_m - 1, 1);      %  ������
bxry = cell(raw2_m - 1, 1);      %  ��д��Ա
jcry = cell(raw2_m - 1, 1);      %  �����Ա
fhry = cell(raw2_m - 1, 1);      %  ����(ǩ��)
shry = cell(raw2_m - 1, 1);      %  ���(ǩ��)
pzry = cell(raw2_m - 1, 1);      %  ��׼(ǩ��)
pzjb = cell(raw2_m - 1, 1);      %  ��׼����
jcdw = cell(raw2_m - 1, 1);      %  ��ⵥλ
dwdz = cell(raw2_m - 1, 1);      %  ��ַ
dwdh = cell(raw2_m - 1, 1);      %  �绰
yzbm = cell(raw2_m - 1, 1);      %  ��������
dwcz = cell(raw2_m - 1, 1);      %  ����
bgrq = cell(raw2_m - 1, 1);      %  ��������
bgys = cell(raw2_m - 1, 1);      %  ����ҳ��
jsdw = cell(raw2_m - 1, 1);      %  ���赥λ
kcdw = cell(raw2_m - 1, 1);      %  ���쵥λ
sjdw = cell(raw2_m - 1, 1);      %  ��Ƶ�λ
zjb = cell(raw2_m - 1, 1);       %  �ܼ��
jldw = cell(raw2_m - 1, 1);      %  ����λ
zdjl = cell(raw2_m - 1, 1);      %  פ�ؼ���
sgdw = cell(raw2_m - 1, 1);      %  ʩ����λ
jcxs = cell(raw2_m - 1, 1);      %  ������ʽ
sbjgxs = cell(raw2_m - 1, 1);    %  �ϲ��ṹ��ʽ
jcxm = cell(raw2_m - 1, 1);      %  �����Ŀ
jcff = cell(raw2_m - 1, 1);      %  ��ⷽ��
jcsl = cell(raw2_m - 1, 1);      %  �������
gcgk = cell(raw2_m - 1, 1);      %  ���̸ſ�
dztj = cell(raw2_m - 1, 1);      %  ��������
sggc = cell(raw2_m - 1, 1);      %  ʩ������
jcgc = cell(raw2_m - 1, 1);      %  ������
zwpmbzt = cell(raw2_m - 1, 1);   %  ׮λƽ�沼��ͼ
bz = cell(raw2_m - 1, 1);        %  ��ע
zysx = cell(raw2_m - 1, 1);      %  ע������

for ii = 1:1:raw1_m                                                        %  ������Ϣ.xlsx ���ݷ���
    switch raw1{ii,1}
        case '���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bgxh{jj} = num2str(raw1{ii , 2});
                else
                    bgxh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bgbh{jj} = num2str(raw1{ii , 2});
                else
                    bgbh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�ṹ������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jgwmc{jj} = num2str(raw1{ii , 2});
                else
                    jgwmc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��̨���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dtbh{jj} = num2str(raw1{ii , 2});
                else
                    dtbh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��׮���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jzbh{jj} = num2str(raw1{ii , 2});
                else
                    jzbh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '������ǿ�ȵȼ�'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    concreteStrength{jj} = num2str(raw1{ii , 2});
                else
                    concreteStrength{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '���׮��(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjzc{jj} = num2str(raw1{ii , 2});
                else
                    sjzc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'ʵ��׮��(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sczc{jj} = num2str(raw1{ii , 2});
                else
                    sczc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '���׮��(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjzj{jj} = num2str(raw1{ii , 2});
                else
                    sjzj{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '���׮�����(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjzTbg{jj} = num2str(raw1{ii , 2});
                else
                    sjzTbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '���׮�˱��(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjzBbg{jj} = num2str(raw1{ii , 2});
                else
                    sjzBbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'ʵ��׮�����(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sczTbg{jj} = num2str(raw1{ii , 2});
                else
                    sczTbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'ʵ��׮�˱��(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sczBbg{jj} = num2str(raw1{ii , 2});
                else
                    sczBbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ע����������(m3)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    gzfl{jj} = num2str(raw1{ii , 2});
                else
                    gzfl{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ע����'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    gzrq{jj} = num2str(raw1{ii , 2});
                else
                    gzrq{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcrq{jj} = num2str(raw1{ii , 2});
                else
                    jcrq{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�����'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcjg{jj} = num2str(raw1{ii , 2});
                else
                    jcjg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'ȱ���������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    qxfx{jj} = num2str(raw1{ii , 2});
                else
                    qxfx{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ܸ���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cggs{jj} = num2str(raw1{ii , 2});
                else
                    cggs{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ܼ��1-2(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj1_2{jj} = num2str(raw1{ii , 2});
                else
                    cgjj1_2{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ܼ��1-3(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj1_3{jj} = num2str(raw1{ii , 2});
                else
                    cgjj1_3{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ܼ��2-3(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj2_3{jj} = num2str(raw1{ii , 2});
                else
                    cgjj2_3{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ܼ��1-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj1_4{jj} = num2str(raw1{ii , 2});
                else
                    cgjj1_4{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ܼ��2-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj2_4{jj} = num2str(raw1{ii , 2});
                else
                    cgjj2_4{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ܼ��3-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj3_4{jj} = num2str(raw1{ii , 2});
                else
                    cgjj3_4{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '����ҳ��'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    fjys{jj} = num2str(raw1{ii , 2});
                else
                    fjys{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�����(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cdjj{jj} = num2str(raw1{ii , 2});
                else
                    cdjj{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�����豸'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    yqsb{jj} = num2str(raw1{ii , 2});
                else
                    yqsb{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '����׼'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcbz{jj} = num2str(raw1{ii , 2});
                else
                    jcbz{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��׼���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jzbg{jj} = num2str(raw1{ii , 2});
                else
                    jzbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    gcmc{jj} = num2str(raw1{ii , 2});
                else
                    gcmc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'ί�е�λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    wtdw{jj} = num2str(raw1{ii , 2});
                else
                    wtdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcnr{jj} = num2str(raw1{ii , 2});
                else
                    jcnr{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jclb{jj} = num2str(raw1{ii , 2});
                else
                    jclb{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��д��Ա'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bxry{jj} = num2str(raw1{ii , 2});
                else
                    bxry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�����Ա'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcry{jj} = num2str(raw1{ii , 2});
                else
                    jcry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '����(ǩ��)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    fhry{jj} = num2str(raw1{ii , 2});
                else
                    fhry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '���(ǩ��)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    shry{jj} = num2str(raw1{ii , 2});
                else
                    shry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��׼(ǩ��)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    pzry{jj} = num2str(raw1{ii , 2});
                else
                    pzry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��׼����'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    pzjb{jj} = num2str(raw1{ii , 2});
                else
                    pzjb{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ⵥλ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcdw{jj} = num2str(raw1{ii , 2});
                else
                    jcdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ַ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dwdz{jj} = num2str(raw1{ii , 2});
                else
                    dwdz{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�绰'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dwdh{jj} = num2str(raw1{ii , 2});
                else
                    dwdh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    yzbm{jj} = num2str(raw1{ii , 2});
                else
                    yzbm{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '����'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dwcz{jj} = num2str(raw1{ii , 2});
                else
                    dwcz{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bgrq{jj} = num2str(raw1{ii , 2});
                else
                    bgrq{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '����ҳ��'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bgys{jj} = num2str(raw1{ii , 2});
                else
                    bgys{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '���赥λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jsdw{jj} = num2str(raw1{ii , 2});
                else
                    jsdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '���쵥λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    kcdw{jj} = num2str(raw1{ii , 2});
                else
                    kcdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��Ƶ�λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjdw{jj} = num2str(raw1{ii , 2});
                else
                    sjdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�ܼ��'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    zjb{jj} = num2str(raw1{ii , 2});
                else
                    zjb{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '����λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jldw{jj} = num2str(raw1{ii , 2});
                else
                    jldw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'פ�ؼ���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    zdjl{jj} = num2str(raw1{ii , 2});
                else
                    zdjl{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'ʩ����λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sgdw{jj} = num2str(raw1{ii , 2});
                else
                    sgdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '������ʽ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcxs{jj} = num2str(raw1{ii , 2});
                else
                    jcxs{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�ϲ��ṹ��ʽ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sbjgxs{jj} = num2str(raw1{ii , 2});
                else
                    sbjgxs{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�����Ŀ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcxm{jj} = num2str(raw1{ii , 2});
                else
                    jcxm{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ⷽ��'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcff{jj} = num2str(raw1{ii , 2});
                else
                    jcff{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '�������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcsl{jj} = num2str(raw1{ii , 2});
                else
                    jcsl{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '���̸ſ�'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    gcgk{jj} = num2str(raw1{ii , 2});
                else
                    gcgk{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dztj{jj} = num2str(raw1{ii , 2});
                else
                    dztj{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'ʩ������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sggc{jj} = num2str(raw1{ii , 2});
                else
                    sggc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcgc{jj} = num2str(raw1{ii , 2});
                else
                    jcgc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '׮λƽ�沼��ͼ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    zwpmbzt{jj} = num2str(raw1{ii , 2});
                else
                    zwpmbzt{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '��ע'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bz{jj} = num2str(raw1{ii , 2});
                else
                    bz{jj} = raw1{ii , 2};
                end
            end
            continue;
        case 'ע������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    zysx{jj} = num2str(raw1{ii , 2});
                else
                    zysx{jj} = raw1{ii , 2};
                end
            end
            continue;
            
        otherwise
            disp('δ���ò�����');
            disp(['    ',raw1{ii,1}]);
    end
end


for ii = 1:1:raw2_n                                                        %  �����Ϣ.xlsx ���ݷ���
    switch raw2{1,ii}
        case '���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bgxh{jj} = num2str(raw2{jj+1 , ii});
                else
                    bgxh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bgbh{jj} = num2str(raw2{jj+1 , ii});
                else
                    bgbh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�ṹ������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jgwmc{jj} = num2str(raw2{jj+1 , ii});
                else
                    jgwmc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��̨���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dtbh{jj} = num2str(raw2{jj+1 , ii});
                else
                    dtbh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��׮���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jzbh{jj} = num2str(raw2{jj+1 , ii});
                else
                    jzbh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '������ǿ�ȵȼ�'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    concreteStrength{jj} = num2str(raw2{jj+1 , ii});
                else
                    concreteStrength{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '���׮��(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjzc{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjzc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'ʵ��׮��(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sczc{jj} = num2str(raw2{jj+1 , ii});
                else
                    sczc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '���׮��(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjzj{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjzj{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '���׮�����(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjzTbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjzTbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '���׮�˱��(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjzBbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjzBbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'ʵ��׮�����(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sczTbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    sczTbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'ʵ��׮�˱��(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sczBbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    sczBbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;   
        case '��ע����������(m3)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    gzfl{jj} = num2str(raw2{jj+1 , ii});
                else
                    gzfl{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ע����'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    gzrq{jj} = num2str(raw2{jj+1 , ii});
                else
                    gzrq{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcrq{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcrq{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�����'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcjg{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcjg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'ȱ���������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    qxfx{jj} = num2str(raw2{jj+1 , ii});
                else
                    qxfx{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ܸ���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cggs{jj} = num2str(raw2{jj+1 , ii});
                else
                    cggs{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ܼ��1-2(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj1_2{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj1_2{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ܼ��1-3(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj1_3{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj1_3{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ܼ��2-3(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj2_3{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj2_3{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ܼ��1-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj1_4{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj1_4{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ܼ��2-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj2_4{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj2_4{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ܼ��3-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj3_4{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj3_4{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '����ҳ��'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    fjys{jj} = num2str(raw2{jj+1 , ii});
                else
                    fjys{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�����(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cdjj{jj} = num2str(raw2{jj+1 , ii});
                else
                    cdjj{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�����豸'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    yqsb{jj} = num2str(raw2{jj+1 , ii});
                else
                    yqsb{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '����׼'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcbz{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcbz{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��׼���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jzbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    jzbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    gcmc{jj} = num2str(raw2{jj+1 , ii});
                else
                    gcmc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'ί�е�λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    wtdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    wtdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcnr{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcnr{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jclb{jj} = num2str(raw2{jj+1 , ii});
                else
                    jclb{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��д��Ա'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bxry{jj} = num2str(raw2{jj+1 , ii});
                else
                    bxry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�����Ա'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcry{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '����(ǩ��)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    fhry{jj} = num2str(raw2{jj+1 , ii});
                else
                    fhry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '���(ǩ��)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    shry{jj} = num2str(raw2{jj+1 , ii});
                else
                    shry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��׼(ǩ��)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    pzry{jj} = num2str(raw2{jj+1 , ii});
                else
                    pzry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��׼����'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    pzjb{jj} = num2str(raw2{jj+1 , ii});
                else
                    pzjb{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ⵥλ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ַ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dwdz{jj} = num2str(raw2{jj+1 , ii});
                else
                    dwdz{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�绰'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dwdh{jj} = num2str(raw2{jj+1 , ii});
                else
                    dwdh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    yzbm{jj} = num2str(raw2{jj+1 , ii});
                else
                    yzbm{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '����'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dwcz{jj} = num2str(raw2{jj+1 , ii});
                else
                    dwcz{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bgrq{jj} = num2str(raw2{jj+1 , ii});
                else
                    bgrq{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '����ҳ��'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bgys{jj} = num2str(raw2{jj+1 , ii});
                else
                    bgys{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '���赥λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jsdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    jsdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '���쵥λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    kcdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    kcdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��Ƶ�λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�ܼ��'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    zjb{jj} = num2str(raw2{jj+1 , ii});
                else
                    zjb{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '����λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jldw{jj} = num2str(raw2{jj+1 , ii});
                else
                    jldw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'פ�ؼ���'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    zdjl{jj} = num2str(raw2{jj+1 , ii});
                else
                    zdjl{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'ʩ����λ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sgdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    sgdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '������ʽ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcxs{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcxs{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�ϲ��ṹ��ʽ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sbjgxs{jj} = num2str(raw2{jj+1 , ii});
                else
                    sbjgxs{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�����Ŀ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcxm{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcxm{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ⷽ��'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcff{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcff{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '�������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcsl{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcsl{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '���̸ſ�'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    gcgk{jj} = num2str(raw2{jj+1 , ii});
                else
                    gcgk{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dztj{jj} = num2str(raw2{jj+1 , ii});
                else
                    dztj{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'ʩ������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sggc{jj} = num2str(raw2{jj+1 , ii});
                else
                    sggc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcgc{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcgc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '׮λƽ�沼��ͼ'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    zwpmbzt{jj} = num2str(raw2{jj+1 , ii});
                else
                    zwpmbzt{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '��ע'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bz{jj} = num2str(raw2{jj+1 , ii});
                else
                    bz{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case 'ע������'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    zysx{jj} = num2str(raw2{jj+1 , ii});
                else
                    zysx{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
            
        otherwise
            disp('δ���ò�����');
            disp(['    ',raw2{1,ii}]);
    end
end


end