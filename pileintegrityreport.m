function [  ] = pileintegrityreport( varargin )
%   PileIntegrityReport  创建桩身完整性超声波透射法检测报告 v20200817
%   将多份报告输出到一个文件中
%   目录下必须有“检测报告”文件夹，“报告中引用文件”中须有相关引用文件存在
%   pileintegrityreport( 2 )   可输入1个正整数，表示输出‘序号’2的检测报告
%   pileintegrityreport( 2, 8 )   可输入2个正整数，表示输出‘序号’2~8的检测报告
%   v20200817  添加引用的文件是否存在的判断


switch nargin
    case 0                                                                  % 不输入参数时执行下面命令
        beginNo1 = input('报告输出开始序号 : ' );    %   键盘输入开始序号
        endNo1 = input('报告输出结束序号 : ' );      %   键盘输入结束序号
    case 1                                                                  %  输入1个参数时执行下面命令，输出一个报告
        beginNo1 = varargin{1};
        endNo1 =  varargin{1};
    case 2                                                                  %   输入2个参数时执行下面命令
        beginNo1 = varargin{1};
        endNo1 =  varargin{2};
    otherwise
        error('More Than 1 Pile Parameters File !');
end

tic;                                                                        %  开始计时

%%  读入参数

disp('读入数据参数信息');

[~,~,raw1] = xlsread('基本信息.xlsx');                                       %  参数文件
[~,~,raw2] = xlsread('检测信息.xlsx');                                       %  参数文件

scgsytspec_3 = [ pwd '\报告中引用文件\3根声测管方位示意图.jpg' ];              %  3根声测管方位示意图存放路径
scgsytspec_4 = [ pwd '\报告中引用文件\4根声测管方位示意图.jpg' ];              %  4根声测管方位示意图存放路径
logospc = [ pwd '\报告中引用文件\河南省交院工程检测科技有限公司徽标.jpg' ];     %  公司logo
filedirspc = [pwd '\检测报告'];                                                 %  检测报告存放文件夹

if ~exist(scgsytspec_3, 'file')
    error(['文件不存在：', scgsytspec_3]);
end
if ~exist(scgsytspec_3, 'file')
    error(['文件不存在：', scgsytspec_4]);
end
if ~exist(scgsytspec_3, 'file')
    error(['文件不存在：', logospc]);
end
if ~exist(filedirspc, 'dir')
    mkdir(filedirspc);
end

%%  输入word文档的基本设置信息

logosize = [ 209.25 , 37.65 ];                                              %  封面logo尺寸，[ 宽 ， 高(磅) ]
exportAsPDF = 1 ;                                                           %  是否输出PDF格式文件   1 - 是；0 - 否
exportFlyPage = 1 ;                                                         %  是否输出扉页“注意事项”   1 - 是；0 - 否
supervisionPattern = 1 ;                                                    %  监理模式选项，  1 - 总监办、驻地监理模式； 2 - 单一监理单位模式

%%  原始数据处理


[ bgxh , bgbh , jgwmc , dtbh , jzbh , concreteStrength , sjzc , sczc , sjzj , gzfl , gzrq , jcrq , jcjg , ...              %    读入参数，参数具体含义见readPara
    qxfx , cggs , cgjj1_2 , cgjj1_3 , cgjj2_3 , cgjj1_4 , cgjj2_4 , cgjj3_4 , fjys , ...
    cdjj , yqsb , jcbz , jzbg , gcmc , wtdw, jcnr, jclb , bxry , jcry , fhry , ...
    shry , pzry , pzjb , jcdw , dwdz , dwdh , yzbm , dwcz , bgrq , bgys , ...
    jsdw , kcdw , sjdw , zjb , jldw , zdjl , sgdw , jcxs , sbjgxs ,  ...
    jcxm , jcff , jcsl , gcgk , dztj , sggc , jcgc , zwpmbzt , bz , ...
    zysx , sjzTbg , sjzBbg , sczTbg , sczBbg ] ...
    = readParameters( raw1 , raw2 );


beginNo = beginNo1 - str2double(bgxh{1}) + 1 ;                              % 输出序号整理为从1开头的数列
endNo = endNo1 - str2double(bgxh{1}) + 1 ;                                  % 输出序号整理为从1开头的数列

zwpmbztspec = fullfile(pwd,'\报告中引用文件\',zwpmbzt);
for ii = beginNo:1:endNo  
    if ~exist(string(zwpmbztspec{ii}), 'file')
        error(['文件不存在：', zwpmbztspec{ii}]);
    end
end

%% 报告封面扉页输出

disp('输出报告封面扉页')
cover( beginNo , endNo , logospc , logosize , exportAsPDF , exportFlyPage ,  ...
    zysx , bgxh , bgbh , gcmc , jcxm , jcff , jcnr , jcdw , bgrq )                           %  报告封面扉页输出函数
disp('报告封面扉页输出完成');

%%

disp('输出报告页')
reportPage( beginNo , endNo , scgsytspec_3 , scgsytspec_4 , exportAsPDF , supervisionPattern , ...                      %   报告页输出函数
    bgxh , bgbh , jgwmc , dtbh , jzbh , concreteStrength , sjzc , sczc , sjzj , gzfl , gzrq , jcrq , jcjg , ...
    qxfx , cggs , cgjj1_2 , cgjj1_3 , cgjj2_3 , cgjj1_4 , cgjj2_4 , cgjj3_4 , fjys , ...
    cdjj , yqsb , jcbz , jzbg , gcmc , wtdw , jcnr , jclb , bxry , jcry , fhry , ...
    shry , pzry , pzjb , jcdw , dwdz , dwdh , yzbm , dwcz , bgrq , bgys , ...
    jsdw , kcdw , sjdw , zjb , jldw , zdjl , sgdw , jcxs , sbjgxs ,  ...
    jcxm , jcff , jcsl , gcgk , dztj , sggc , jcgc , zwpmbzt , bz , ...
    zysx , sjzTbg , sjzBbg , sczTbg , sczBbg );
disp('报告页输出完成');


%%

disp('报告输出完成');
toc;                                                                       %  结束计时
disp(' ');

% reply = input('Do you want more? Y/N [Y]: ', 's');
% if isempty(reply)
%     reply = 'Y';
% end

end


function [] = cover( beginNo , endNo , logospc , logosize , exportAsPDF , exportFlyPage , ...
    zysx , bgxh , bgbh , gcmc , jcxm , jcff , jcnr , jcdw , bgrq )
%  cover  输出报告封面YJS-150  v20171122


%%  输入word文档的基本设置信息

paperSize = 'wdPaperA4';                                                    %  设置纸张大小    'wdPaperA4'     'wdPaperA3'   'wdPaperB5' ......
orientation = 'wdOrientPortrait';                                           %  设置纸张方向   'wdOrientPortrait'  'wdOrientLandscape'
margin = [42.55, 42.55, 56.7, 56.7];                                      %  设置页边距，单位point(磅)    [TopMargin, BottomMargin, LeftMargin, RightMargin]
headerFooterDistance = [42.55, 42.55];                                       %  设置页眉页脚距边界，单位point(磅)   [页眉，页脚]


%%  调用Microsoft Word 服务器

disp('    调用 Microsoft Word 服务器');

try                                                                         %  调用actxserver函数创建Microsoft Word 服务器
    word = actxGetRunningServer('Word.Application');                        %  若word服务器已经打开，返回其句柄Word
catch
    word = actxserver('Word.Application');                                  %  若word服务器没有打开，创建一个，返回句柄Word
end

set(word,'Visible',0);                                                      %  设置服务器界面变为可见状态, 1 - 可见； 0 - 不可见
document = invoke(word.Documents,'Add');                                    %  打开一个空白文档

selection = word.Selection;                                                 %  返回Word.Selection接口句柄
leftTabstop = margin(3) + 135;                                              %  设置页眉右端对齐制表符位置

%%  页面设置

disp('    文档页面设置');
document.PageSetup.PaperSize = paperSize;                                   % 设置纸张大小，A4、A3、B5
document.PageSetup.Orientation = orientation;                               % 设置纸张方向
document.PageSetup.TopMargin = margin(1);                                   % 页面边距设置
document.PageSetup.BottomMargin = margin(2);
document.PageSetup.LeftMargin = margin(3);
document.PageSetup.RightMargin = margin(4);
document.PageSetup.HeaderDistance = headerFooterDistance(1);                %  设置页眉距边界距离
document.PageSetup.FooterDistance = headerFooterDistance(2);                %  设置页脚距边界距离

%% 样式设置

disp('    文档样式设置');


try
    userStyles = document.Styles.Item('正文');
catch
    userStyles = document.Styles.Add('正文');
end                                                                         %  设置“正文”样式
userStyles.Font.Name = '仿宋_GB2312';                                       %  设置默认字体为宋体
userStyles.Font.NameFarEast = '仿宋_GB2312';                                %  设置中文字体为宋体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 14.0;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';           %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)  负值不是悬挂
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  设置行距    'wdLineSpaceSingle' 'wdLineSpace1pt5'

%%  输出报告封面扉页

for ii = beginNo : 1 : endNo
    
    disp(['    输出序号 ', bgxh{ii} ,' 报告编号( ' , bgbh{ii} ,' )封面']);
    
    if ii ~= beginNo
        selection.InsertBreak(2);                                           %  插入分页符0或分节符2
    end
    
    selection.Style = '正文';
    selection.ParagraphFormat.SpaceAfter = 42;                              %  段后间距(磅)
    selection.TypeParagraph;
    selection.Text = strcat('报告编号：', bgbh{ii} );                        %  第二行，报告编号
    selection.Font.Name = '新宋体';                                          %  设置默认字体为宋体
    selection.Font.NameFarEast = '新宋体';                                   %  设置中文字体为宋体
    selection.Font.NameAscii = 'Times New Roman';                            %  设置Ascii字体为Times New Roman
    selection.Font.NameOther = 'Times New Roman';                            %  设置其他字符字体为Times New Roman
    selection.Font.Size = 12.0;                                              %  设置字体大小(磅)
    selection.Font.Bold = 1;                                                 %  设置字体加粗
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphRight';           %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
    selection.ParagraphFormat.SpaceBefore = 6;                               %  段前间距(磅)
    selection.ParagraphFormat.SpaceAfter = 12;                               %  段后间距(磅)
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.Text = gcmc{ii};                                               %
    selection.Font.Name = '华文中宋';                                         %  设置默认字体为宋体
    selection.Font.NameFarEast = '华文中宋';                                  %  设置中文字体为宋体
    selection.Font.NameAscii = 'Times New Roman';                             %  设置Ascii字体为Times New Roman
    selection.Font.NameOther = 'Times New Roman';                             %  设置其他字符字体为Times New Roman
    selection.Font.Size = 18.0;                                               %  设置字体大小(磅)
    selection.Font.Bold = 1;                                                  %  设置字体加粗
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';           %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
    selection.ParagraphFormat.SpaceBefore = 0;                                %  段前间距(磅)
    selection.ParagraphFormat.SpaceAfter = 0;                                 %  段后间距(磅)
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.Font.Size = 18.0;                                               %  设置字体大小(磅)
    
    selection.TypeParagraph;
    selection.Text = strcat( '混凝土灌注桩桩身完整性',  11 , '超声波透射法检测报告');                %  中间主标题   selection.Text = strcat( '混凝土灌注桩桩身完整性',  11 , '超声波透射法检测报告');    selection.Text = strcat( jcxm{ii}, jcff{ii}, '检测报告');
    selection.Font.Name = '华文中宋';                                        %  设置默认字体为宋体
    selection.Font.NameFarEast = '华文中宋';                                 %  设置中文字体为宋体
    selection.Font.NameAscii = 'Times New Roman';                            %  设置Ascii字体为Times New Roman
    selection.Font.NameOther = 'Times New Roman';                            %  设置其他字符字体为Times New Roman
    selection.Font.Size = 36.0;                                              %  设置字体大小(磅)
    selection.Font.Bold = 1;                                                 %  设置字体加粗
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';          %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.ParagraphFormat.SpaceAfter = 222;                              %  段后间距(磅)
    
    selection.TypeParagraph;                                                 %  回车，插入一行
    selection.Style = '正文';
    handle1 = selection.InlineShape.AddPicture(logospc);                     % 插入公司logo
    handle1.Width = logosize(1);                                             % logo宽、高
    handle1.Height = logosize(2);
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.Text = jcdw{ii};
    selection.Font.Name = '仿宋_GB2312';                                      %  设置默认字体为宋体
    selection.Font.NameFarEast = '仿宋_GB2312';                               %  设置中文字体为宋体
    selection.Font.NameAscii = 'Times New Roman';                             %  设置Ascii字体为Times New Roman
    selection.Font.NameOther = 'Times New Roman';                             %  设置其他字符字体为Times New Roman
    selection.Font.Size = 18.0;                                               %  设置字体大小(磅)
    selection.Font.Bold = 0;                                                  %  设置字体加粗
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';           %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
    selection.MoveDown;
    
    selection.TypeParagraph;
%     yyyy = num2str( year( bgrq{ii} ) );                                  %  提取年月日
%     mm = num2str( month( bgrq{ii} ) );
%     dd = num2str( day( bgrq{ii} ) );
    yyyy = num2str( datestr( strtok(bgrq{ii}) , 10 ) );                            %  采用datestr函数提取年月日
    mm = num2str( datestr( strtok(bgrq{ii}) , 5 ) );
    dd = num2str( datestr( strtok(bgrq{ii}) , 7 ) );
    
    yyyy = strrep(yyyy, '0', '〇');
    yyyy = strrep(yyyy, '1', '一');
    yyyy = strrep(yyyy, '2', '二');
    yyyy = strrep(yyyy, '3', '三');
    yyyy = strrep(yyyy, '4', '四');
    yyyy = strrep(yyyy, '5', '五');
    yyyy = strrep(yyyy, '6', '六');
    yyyy = strrep(yyyy, '7', '七');
    yyyy = strrep(yyyy, '8', '八');
    yyyy = strrep(yyyy, '9', '九');
  
    mm = strrep(mm, '10', '十');
    mm = strrep(mm, '11', '十一');
    mm = strrep(mm, '12', '十二');
    mm = strrep(mm, '01', '一');
    mm = strrep(mm, '02', '二');
    mm = strrep(mm, '03', '三');
    mm = strrep(mm, '04', '四');
    mm = strrep(mm, '05', '五');
    mm = strrep(mm, '06', '六');
    mm = strrep(mm, '07', '七');
    mm = strrep(mm, '08', '八');
    mm = strrep(mm, '09', '九');

    dd = strrep(dd, '0', '〇');
    dd = strrep(dd, '1', '一');
    dd = strrep(dd, '2', '二');
    dd = strrep(dd, '3', '三');
    dd = strrep(dd, '4', '四');
    dd = strrep(dd, '5', '五');
    dd = strrep(dd, '6', '六');
    dd = strrep(dd, '7', '七');
    dd = strrep(dd, '8', '八');
    dd = strrep(dd, '9', '九');
    selection.Text = strcat( yyyy , '年' , mm , '月');
    selection.Font.Name = '宋体';                                              %  设置默认字体为宋体
    selection.Font.NameFarEast = '宋体';                                       %  设置中文字体为宋体
    selection.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
    selection.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
    selection.Font.Size = 18.0;                                                %  设置字体大小(磅)
    selection.Font.Bold = 0;                                                   %  设置字体加粗
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
    selection.MoveDown;
    
    selection.TypeParagraph;
    selection.TypeParagraph;
    selection.Text = '表YJS-1150-F/0';
    selection.Font.Name = '新宋体';                                            %  设置默认字体为宋体
    selection.Font.NameFarEast = '新宋体';                                     %  设置中文字体为宋体
    selection.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
    selection.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
    selection.Font.Size = 10.5;                                                %  设置字体大小(磅)
    selection.Font.Bold = 0;                                                   %  设置字体加粗
    selection.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';              %  设置对齐方式 'wdAlignParagraph
    selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
    selection.MoveDown;
    
    if exportFlyPage == 1
        disp(['    输出序号 ', bgxh{ii} ,' 报告编号( ' , bgbh{ii} ,' )扉页“注意事项”']);
        selection.InsertBreak(0);                                             %  插入分页符0或分节符2
        selection.Style = '正文';
        selection.ParagraphFormat.SpaceAfter = 28;                              %  段后间距(磅)
        selection.TypeParagraph;
        selection.Text = '注 意 事 项';                                       %  标题
        selection.Font.Name = '宋体';                                         %  设置默认字体为宋体
        selection.Font.NameFarEast = '宋体';                                  %  设置中文字体为宋体
        selection.Font.NameAscii = 'Times New Roman';                         %  设置Ascii字体为Times New Roman
        selection.Font.NameOther = 'Times New Roman';                         %  设置其他字符字体为Times New Roman
        selection.Font.Size = 28.0;                                           %  设置字体大小(磅)
        selection.Font.Bold = 1;                                              %  设置字体加粗
        selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';       %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
        selection.ParagraphFormat.SpaceBefore = 0;                            %  段前间距(磅)
        selection.ParagraphFormat.SpaceAfter = 17;                            %  段后间距(磅)
        selection.MoveDown;
        
        selection.TypeParagraph;
        selection.Text = zysx{ii};                                            %  "注意事项"内容
        selection.Font.Name = '仿宋_GB2312';                                  %  设置默认字体为宋体
        selection.Font.NameFarEast = '仿宋_GB2312';                           %  设置中文字体为宋体
        selection.Font.NameAscii = 'Times New Roman';                         %  设置Ascii字体为Times New Roman
        selection.Font.NameOther = 'Times New Roman';                         %  设置其他字符字体为Times New Roman
        selection.Font.Size = 14.0;                                           %  设置字体大小(磅)
        selection.Font.Bold = 0;                                              %  设置字体加粗
        selection.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';      %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
        selection.ParagraphFormat.CharacterUnitFirstLineIndent = -2;          %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
        selection.ParagraphFormat.SpaceBefore = 0;                            %  段前间距(磅)
        selection.ParagraphFormat.SpaceAfter = 0;                             %  段后间距(磅)
        selection.MoveDown;
        
        selection.TypeParagraph;
        selection.ParagraphFormat.SpaceAfter = 42;                                  %  段后间距(磅)
        
        selection.TypeParagraph;
        selection.Text = strcat( 09 , 09 , jcdw{ii});
        selection.Font.Name = '仿宋_GB2312';                                   %  设置默认字体为宋体
        selection.Font.NameFarEast = '仿宋_GB2312';                            %  设置中文字体为宋体
        selection.Font.NameAscii = 'Times New Roman';                          %  设置Ascii字体为Times New Roman
        selection.Font.NameOther = 'Times New Roman';                          %  设置其他字符字体为Times New Roman
        selection.Font.Size = 14.0;                                            %  设置字体大小(磅)
        selection.Font.Bold = 0;                                               %  设置字体加粗
        selection.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';       %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
        selection.ParagraphFormat.SpaceBefore = 0;                             %  段前间距(磅)
        selection.ParagraphFormat.SpaceAfter = 0;                              %  段后间距(磅)
        tabstop = selection.ParagraphFormat.TabStops.Add( leftTabstop );       %  设置第二段制表符位置
        tabstop.Alignment = 'wdAlignTabLeft';
        selection.MoveDown;
        selection.TypeParagraph;
    elseif exportFlyPage == 0
        disp('    不输出扉页“注意事项”');
    else
        error('输出扉页“注意事项” 选择错误');
    end
    
    
end

%%%%%%  保存封面扉页文件  %%%%%%
disp('    保存封面扉页文档');
bgmc = strcat( '封面扉页', bgbh{beginNo}, '～' , bgbh{endNo}  );              %  输出检测报告名称
filespec= fullfile( pwd , '\检测报告\' , bgmc );                              %  输出报告文件存放路径
try
    document.SaveAs(filespec);
catch
    document.SaveAs2(filespec);
end

if exportAsPDF == 1
    if  word.Version >= 12
        disp('    输出PDF格式文件');
        document.ExportAsFixedFormat( filespec , 17 );                        %  另存为PDF格式   17 - PDF格式，18 - XPS格式
    end
end

document.Close;
% word.Quit;

end


function [] = reportPage( beginNo , endNo , scgsytspec_3 , scgsytspec_4 , exportAsPDF , supervisionPattern , ...     %   报告页输出函数
    bgxh , bgbh , jgwmc , dtbh , jzbh , concreteStrength , sjzc , sczc , sjzj , gzfl , gzrq , jcrq , jcjg , ...
    qxfx , cggs , cgjj1_2 , cgjj1_3 , cgjj2_3 , cgjj1_4 , cgjj2_4 , cgjj3_4 , fjys , ...
    cdjj , yqsb , jcbz , jzbg , gcmc , wtdw , jcnr , jclb , bxry , jcry , fhry , ...
    shry , pzry , pzjb , jcdw , dwdz , dwdh , yzbm , dwcz , bgrq , bgys , ...
    jsdw , kcdw , sjdw , zjb , jldw , zdjl , sgdw , jcxs , sbjgxs ,  ...
    jcxm , jcff , jcsl , gcgk , dztj , sggc , jcgc , zwpmbzt , bz , ...
    zysx , sjzTbg , sjzBbg , sczTbg , sczBbg )
%  reportPage 报告正文页输出    v20180606


%%  输入word文档的基本设置信息

paperSize = 'wdPaperA4';                                                    %  设置纸张大小    'wdPaperA4'     'wdPaperA3'   'wdPaperB5' ......
orientation = 'wdOrientPortrait';                                           %  设置纸张方向   'wdOrientPortrait'  'wdOrientLandscape'
margin = [65.2, 53.85, 53.85, 53.85];                                        %  设置页边距，单位point(磅)    [TopMargin, BottomMargin, LeftMargin, RightMargin]
headerFooterDistance = [42.55, 49.6];                                       %  设置页眉页脚距边界，单位point(磅)   [页眉，页脚]
columWidth_1 = [127.6, 90.6, 90.6, 127.6];                                  %  签字页单元格宽度(磅)
rowHeight_1 = [41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41, 41];  %  签字页单元格高度(磅)
height_2 = 28.4;                                                            %  表1  工程简介表每行高度(磅)
width_2 = [111.4 , 331.6];                                                  %  表1  工程简介表每列宽度(磅)
height_3 = [40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40 ];          %  表2  检测情况一览表每行高度(磅)
width_3 = [100 , 100, 100, 100, 100];                                       %  表2  检测情况一览表每列宽度(磅)
picsize_1 = [487.7 , 304.8 , 1];                                            %  图1  桩位平面布置图图片尺寸限值（487.7，304.8），  [ 宽，高(磅)，是否保持原图片宽高比(0/1保持) ]
picsize_2 = [ 90 , 70 ];                                                    %  声测管位置示意图 图片尺寸，    [ 宽 ， 高(磅) ]



%%  调用Microsoft Word 服务器

disp('    调用 Microsoft Word 服务器');

try                                                                         %  调用actxserver函数创建Microsoft Word 服务器
    word = actxGetRunningServer('Word.Application');                        %  若word服务器已经打开，返回其句柄Word
catch
    word = actxserver('Word.Application');                                  %  若word服务器没有打开，创建一个，返回句柄Word
end

set(word,'Visible',0);                                                      %  设置服务器界面变为可见状态, 1 - 可见； 0 - 不可见
document = invoke(word.Documents,'Add');                                    %  打开一个空白文档

% if exist(filespec,'file')
%     document = invoke(word.Documents,'Open',filespec);                    %  若文件存在，打开该文件，否则新建一个文件
% else
%     document = invoke(word.Documents,'Add');                              %  若文件不存在，新建一个文件
%     document.SaveAs2(filespec);                                           %  保存文档
% end

selection = word.Selection;                                                 %  返回Word.Selection接口句柄
content = document.Content;                                                 %  返回Word.Documents.Content接口句柄


%%  页面设置

disp('    文档页面设置');

document.PageSetup.PaperSize = paperSize;                                   % 设置纸张大小，A4、A3、B5
document.PageSetup.LayoutMode = 'wdLayoutModeDefault';                      % 设置文档网格为无网格
document.PageSetup.Orientation = orientation;                               % 设置纸张方向
document.PageSetup.TopMargin = margin(1);                                   % 页面边距设置
document.PageSetup.BottomMargin = margin(2);
document.PageSetup.LeftMargin = margin(3);
document.PageSetup.RightMargin = margin(4);
document.PageSetup.HeaderDistance = headerFooterDistance(1);                %  设置页眉距边界距离
document.PageSetup.FooterDistance = headerFooterDistance(2);                %  设置页脚距边界距离
headerFooterTabstop = document.PageSetup.PageWidth - margin(3) - margin(4); %  设置页眉右端对齐制表符位置


%% 样式设置

disp('    文档样式设置');

try
    userStyles = document.Styles.Item('正文');
catch
    userStyles = document.Styles.Add('正文');
end                                                                        %  设置“正文”样式
userStyles.Font.Name = '宋体';                                              %  设置默认字体为宋体
userStyles.Font.NameFarEast = '宋体';                                       %  设置中文字体为宋体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 12.0;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphJustify';           %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
% userStyles.ParagraphFormat.LeftIndent = 0;                                %  设置缩进：左侧(磅)
% userStyles.ParagraphFormat.FirstlineIndent = 24;                          %  首行缩进(磅)  负值不是悬挂
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 2;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';             %  设置行距    'wdLineSpaceSingle' 'wdLineSpace1pt5'



try
    userStyles = document.Styles.Add('主标题');
catch
    userStyles = document.Styles.Item('主标题');
end                                                                         %  设置“主标题”样式
userStyles.Font.Name = '黑体';                                              %  设置默认字体为黑体
userStyles.Font.NameFarEast = '黑体';                                       %  设置中文字体为黑体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 16.0;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  设置对齐方式'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 8.0;                               %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 8.0;                                %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';             %  设置行距    'wdLineSpaceSingle' 'wdLineSpace1p


try
    userStyles = document.Styles.Add('标题 1');
catch
    userStyles = document.Styles.Item('标题 1');
end                                                                         %  设置“标题 2”样式
userStyles.Font.Name = '黑体';                                              %  设置默认字体为黑体
userStyles.Font.NameFarEast = '黑体';                                       %  设置中文字体为黑体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 15.0;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';              %  设置对齐方式'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevel1';                %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 7.5;                               %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';             %  设置行距    'wdLineSpaceSingle' 'wdLineSpace1pt5'


try
    userStyles = document.Styles.Add('标题 2');
catch
    userStyles = document.Styles.Item('标题 2');
end                                                                         %  设置“标题 2”样式
userStyles.Font.Name = '黑体';                                              %  设置默认字体为黑体
userStyles.Font.NameFarEast = '黑体';                                       %  设置中文字体为黑体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 14.0;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';              %  设置对齐方式'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel ='wdOutlineLevel2';                 %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 1;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';


try
    userStyles = document.Styles.Add('题注');
catch
    userStyles = document.Styles.Item('题注');
end                                                                         %  设置“题注”样式
userStyles.Font.Name = '黑体';                                              %  设置默认字体为黑体
userStyles.Font.NameFarEast = '黑体';                                       %  设置中文字体为黑体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 10.5;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  设置对齐方式'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpace1pt5';


try
    userStyles = document.Styles.Add('表格');
catch
    userStyles = document.Styles.Item('表格');
end                                                                         %  设置“表格”样式
userStyles.Font.Name = '宋体';                                              %  设置默认字体为黑体
userStyles.Font.NameFarEast = '宋体';                                       %  设置中文字体为黑体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 12.0;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  设置对齐方式'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  设置行距    'wdLineSpaceSingle' 'wdLineSpace1pt5'


try
    userStyles = document.Styles.Add('签字页');
catch
    userStyles = document.Styles.Item('签字页');
end                                                                         %  设置“签字页”样式
userStyles.Font.Name = '宋体';                                              %  设置默认字体为黑体
userStyles.Font.NameFarEast = '宋体';                                       %  设置中文字体为黑体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 14.0;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  设置对齐方式'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  设置行距    'wdLineSpaceSingle' 'wdLineSpace1pt5'


try
    userStyles = document.Styles.Item('页眉');
catch
    userStyles = document.Styles.Add('页眉');
end                                                                         %  设置“页眉”样式
userStyles.Font.Name = '宋体';                                              %  设置默认字体为宋体
userStyles.Font.NameFarEast = '宋体';                                       %  设置中文字体为宋体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 10.5;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';              %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)  负值不是悬挂
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  设置行距    'wdLineSpaceSingle' 'wdLineSpace1pt5'
userStyles.ParagraphFormat.TabStops.ClearAll;


try
    userStyles = document.Styles.Item('页脚');
catch
    userStyles = document.Styles.Add('页脚');
end                                                                         %  设置“页脚”样式
userStyles.Font.Name = '宋体';                                              %  设置默认字体为宋体
userStyles.Font.NameFarEast = '宋体';                                       %  设置中文字体为宋体
userStyles.Font.NameAscii = 'Times New Roman';                              %  设置Ascii字体为Times New Roman
userStyles.Font.NameOther = 'Times New Roman';                              %  设置其他字符字体为Times New Roman
userStyles.Font.Size = 10.5;                                                %  设置字体大小(磅)
userStyles.Font.Bold = 0;                                                   %  设置字体加粗
userStyles.Font.Italic = 0;                                                 %  设置字体斜体
userStyles.ParagraphFormat.Alignment = 'wdAlignParagraphCenter';            %  设置对齐方式 'wdAlignParagraphLeft'  'wdAlignParagraphCenter'  'wdAlignParagraphRight' 'wdAlignParagraphJustify'
userStyles.ParagraphFormat.OutlineLevel = 'wdOutlineLevelBodyText';         %  设置大纲级别'wdOutlineLevelBodyText'、'wdOutlineLevel1'
userStyles.ParagraphFormat.LeftIndent = 0;                                  %  设置缩进：左侧(磅)
userStyles.ParagraphFormat.FirstlineIndent = 0;                             %  首行缩进(磅)  负值不是悬挂
userStyles.ParagraphFormat.CharacterUnitLeftIndent = 0;                     %  设置缩进：左侧(字符)
userStyles.ParagraphFormat.CharacterUnitFirstLineIndent = 0;                %  首行缩进(字符)，正为首行缩进，负为悬挂缩进
userStyles.ParagraphFormat.SpaceBefore = 0;                                 %  段前间距(磅)
userStyles.ParagraphFormat.SpaceAfter = 0;                                  %  段后间距(磅)
userStyles.ParagraphFormat.LineSpacingRule = 'wdLineSpaceSingle';           %  设置行距    'wdLineSpaceSingle' 'wdLineSpace1pt5'



%%  输出报告

numTables = 0;                                                              %  表格计数
selection.Style = '正文';
for ii = beginNo : 1 : endNo
    
    disp(['    输出序号 ', bgxh{ii} ,' 报告编号( ' , bgbh{ii},' )报告页']);
    if  bz{ii} ~= char(32)                                                      %  如果备注非空格，屏幕上输出“备注”内容
        disp( [ '              备注：'  , bz{ii} ] );
    end
    
    %%%%%%%%%%%%%%%   签字页 输出   %%%%%%%%%%%%%%%%%%
    
    
    if ii ~= beginNo
        selection.InsertBreak(2);                                           %  插入分页符0或分节符2
    end
    
    selection.TypeParagraph;                                                %  回车，插入一行
    selection.Start = content.end;                                          %  将选定区域的起始位置定位到文章末尾
    document.Tables.Add(selection.Range,16,4);                              %  插入第一页表格
    numTables = numTables + 1;
    DTI = document.Tables.Item( numTables );                                %  获取新建表格句柄
    
    DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';                     %  设置外边框线型，查询用DTI.Borders.set('OutsideLineStyle')
    DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';                      %  设置外边框线宽
    DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';
    DTI.Borders.InsideLineWidth = 'wdLineWidth050pt';
    DTI.Rows.Alignment = 'wdAlignRowCenter';                                %  设置整个表格的水平对齐方式，查询用set函数
    
    
    for mm = 1:1:16
        DTI.Rows.Item(mm).Height = rowHeight_1(mm);                         %  设置签字页单元格高度
        for nn = 1:1:4
            DTI.Columns.Item(nn).Width = columWidth_1(nn);                  %  设置签字页单元格宽度
            DTI.Cell(mm,nn).Range.Paragraphs.Alignment = 'wdAlignParagraphCenter';  %  设置单元格水平居中对齐
            DTI.Cell(mm,nn).VerticalAlignment = 'wdCellAlignVerticalCenter';   %  设置单元格对齐方式
            DTI.Cell(mm,nn).Range.ParagraphFormat.Style = '签字页';          %   设置样式为‘签字页’
            
        end
    end
    
    
    DTI.Cell(1,2).Merge(DTI.Cell(1,4));                                     %  合并单元格，第1行2～4合并
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
    
    
    DTI.Cell(1,1).Range.Text = '报 告 编 号';                                %  录入表格文字
    DTI.Cell(2,1).Range.Text = '工 程 名 称';
    DTI.Cell(3,1).Range.Text = '委 托 单 位';
    DTI.Cell(4,1).Range.Text = '检 测 内 容';
    DTI.Cell(5,1).Range.Text = '检 测 类 别';
    DTI.Cell(6,1).Range.Text = '检 测 人 员';
    DTI.Cell(7,1).Range.Text = '编 写 人 员';    
    DTI.Cell(8,1).Range.Text = '复 核 (签名)';
    DTI.Cell(9,1).Range.Text = '审 核 (签名)';
    DTI.Cell(10,1).Range.Text = '批 准 (签名)';
    DTI.Cell(11,1).Range.Text = '检 测 单 位';
    DTI.Cell(12,1).Range.Text = '地      址';
    DTI.Cell(13,1).Range.Text = '电      话';
    DTI.Cell(14,1).Range.Text = '邮 政 编 码';
    DTI.Cell(14,3).Range.Text = '传  真';
    
    for mm = 1:14                                                           %  字体加黑
        DTI.Cell(mm,1).Range.Font.Bold = 1;
    end
    DTI.Cell(14,3).Range.Font.Bold = 1;
    
    DTI.Cell(1,2).Range.Text = bgbh{ii};      %  报告编号
    DTI.Cell(2,2).Range.Text = gcmc{ii};      %  工程名称
    DTI.Cell(3,2).Range.Text = wtdw{ii};      %  委托单位
    DTI.Cell(4,2).Range.Text = jcnr{ii};      %  检测内容
    DTI.Cell(5,2).Range.Text = jclb{ii};      %  检测类别
    DTI.Cell(6,2).Range.Text = jcry{ii};      %  检测人员
    DTI.Cell(7,2).Range.Text = bxry{ii};      %  编写人员    
    DTI.Cell(8,2).Range.Text = fhry{ii};      %  复核人员
    DTI.Cell(9,2).Range.Text = shry{ii};      %  审核人员
    DTI.Cell(10,2).Range.Text = pzry{ii};     %  批准人员
    DTI.Cell(11,2).Range.Text = jcdw{ii};     %  检测单位
    DTI.Cell(12,2).Range.Text = dwdz{ii};     %  单位地址
    DTI.Cell(13,2).Range.Text = dwdh{ii};     %  单位电话
    DTI.Cell(14,2).Range.Text = yzbm{ii};     %  邮政编号
    DTI.Cell(14,4).Range.Text = dwcz{ii};     %  单位传真
    
    selection.Tables.Item( 1 ).Cell(15,1).Select();
    selection.Text = '报告日期：';
    selection.Font.Bold = 1;
    selection.EndKey();
    selection.Text =  datestr( datenum( strtok(bgrq{ii}) ), 'yyyy年mm月dd日' );
    selection.Font.Bold = 0;
    
    selection.Tables.Item( 1 ).Cell(16,1).Select();
    selection.Text = '报告页：';
    selection.Font.Bold = 1;
    selection.EndKey();
    selection.Text = strcat('第1页 共', bgys{ii}, '页');
    selection.Font.Bold = 0;
    
    selection.Start = content.end;                                          %  将选定区域的起始位置定位到文章末尾
    selection.InsertBreak(2);                                               %  插入分页符0或分节符2
    
    
    %%%%%%%%%%%%%%%%%   签字页 输出完成   %%%%%%%%%%%%%%%%%%
    
    
    
    
    %%%%%%%%%%%%%%%%%        正文页      %%%%%%%%%%%%%%%%%%%%%
    
    selection.Start = content.end;                                         %  将选定区域的起始位置定位到文章末尾
    selection.Text = strcat(jcnr{ii} , '报告')';                           %  正文主标题  %  selection.Text = '混凝土灌注桩桩身完整性超声波透射法检测报告';
    selection.Style = '主标题';                                            %   样式'主标题'
    selection.MoveDown;                                                    %  将光标移到所选区域的最后,段落标志的后面，也即下一段的最前面;区别于EndKey，段落的最后面，段落标志的前面
    
    
    selection.TypeParagraph;                                               %  回车，插入一行
    selection.Text = '1工程简介';
    selection.Style = '标题 1';
    selection.MoveDown;                                                    %  将光标移到所选区域的最后
    
    
    selection.TypeParagraph;                                               %  回车，插入一行
    selection.Text = '表 1  工程简介表';
    selection.Style = '题注';
    selection.MoveDown;                                                    %  将光标移到所选区域的最后
    
    selection.TypeParagraph;
    if supervisionPattern == 1                                             %  根据监理模式选择工程简介表样式； 1 - 总监办、驻地监理模式； 2 - 单一监理单位模式
        document.Tables.Add(selection.Range,12,2);                              %  插入第一页表格
        numTables = numTables + 1;
        DTI = document.Tables.Item( numTables );
        
        for mm = 1:1:12
            DTI.Rows.Item(mm).Height = height_2;                                %  设置表1单元格高度
            for nn = 1:1:2
                DTI.Columns.Item(nn).Width = width_2(nn);                       %  设置表1单元格宽度
                DTI.Cell(mm,nn).Range.Paragraphs.Alignment = 'wdAlignParagraphCenter';  %  设置单元格水平居中对齐
                DTI.Cell(mm,nn).VerticalAlignment = 'wdCellAlignVerticalCenter';        %  设置单元格对齐方式
                DTI.Cell(mm,nn).Range.ParagraphFormat.Style = '表格';            %   设置样式为‘表格’
            end
        end
        
        DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';                     %  设置外边框线型，查询用DTI.Borders.set('OutsideLineStyle')
        DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';                      %  设置外边框线宽
        DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';
        DTI.Borders.InsideLineWidth = 'wdLineWidth050pt';
        DTI.Rows.Alignment = 'wdAlignRowCenter';                                %  设置单元格水平居中对齐
        DTI.Cell(1,1).Range.Text = '工程名称';                                   %  录入表格文字
        DTI.Cell(2,1).Range.Text = '建设单位';
        DTI.Cell(3,1).Range.Text = '勘察单位';
        DTI.Cell(4,1).Range.Text = '设计单位';
        DTI.Cell(5,1).Range.Text = '总监办';
        DTI.Cell(6,1).Range.Text = '驻地监理';
        DTI.Cell(7,1).Range.Text = '施工单位';
        DTI.Cell(8,1).Range.Text = '基础型式';
        DTI.Cell(9,1).Range.Text = '上部结构型式';
        DTI.Cell(10,1).Range.Text = '检测项目';
        DTI.Cell(11,1).Range.Text = '检测方法';
        DTI.Cell(12,1).Range.Text = '检测数量';
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
        
        
    elseif supervisionPattern == 2                                         %  根据监理模式选择工程简介表样式； 1 - 总监办、驻地监理模式； 2 - 单一监理单位模式
        document.Tables.Add(selection.Range,11,2);                              %  插入第一页表格
        numTables = numTables + 1;
        DTI = document.Tables.Item( numTables );
        
        for mm = 1:1:11
            DTI.Rows.Item(mm).Height = height_2;                                %  设置表1单元格高度
            for nn = 1:1:2
                DTI.Columns.Item(nn).Width = width_2(nn);                       %  设置表1单元格宽度
                DTI.Cell(mm,nn).Range.Paragraphs.Alignment = 'wdAlignParagraphCenter';  %  设置单元格水平居中对齐
                DTI.Cell(mm,nn).VerticalAlignment = 'wdCellAlignVerticalCenter';        %  设置单元格对齐方式
                DTI.Cell(mm,nn).Range.ParagraphFormat.Style = '表格';            %   设置样式为‘表格’
            end
        end
        
        DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';                     %  设置外边框线型，查询用DTI.Borders.set('OutsideLineStyle')
        DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';                      %  设置外边框线宽
        DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';
        DTI.Borders.InsideLineWidth = 'wdLineWidth050pt';
        DTI.Rows.Alignment = 'wdAlignRowCenter';                                %  设置单元格水平居中对齐
        DTI.Cell(1,1).Range.Text = '工程名称';                                   %  录入表格文字
        DTI.Cell(2,1).Range.Text = '建设单位';
        DTI.Cell(3,1).Range.Text = '勘察单位';
        DTI.Cell(4,1).Range.Text = '设计单位';
        DTI.Cell(5,1).Range.Text = '监理单位';
        DTI.Cell(6,1).Range.Text = '施工单位';
        DTI.Cell(7,1).Range.Text = '基础型式';
        DTI.Cell(8,1).Range.Text = '上部结构型式';
        DTI.Cell(9,1).Range.Text = '检测项目';
        DTI.Cell(10,1).Range.Text = '检测方法';
        DTI.Cell(11,1).Range.Text = '检测数量';
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
    
    
    
    selection.Text = '2工程概况及地质条件描述';
    selection.Style = '标题 1';
    selection.MoveDown;
    
    
    selection.TypeParagraph;                                                %  回车，插入一行
    selection.Text = '2.1工程概况';
    selection.Style = '标题 2';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  回车，插入一行
    gcgk1 = strrep(gcgk{ii}, 'wtdw', wtdw{ii});                             %  将文中'wtdw'替换为委托单位wtdw{ii}
    gcgk2 = strrep(gcgk1, 'jcdw' , jcdw{ii} );                              %  将文中'jcdw'替换为检测单位jcdw{ii}
    gcgk3 = strrep(gcgk2, 'gcmc' , erase(gcmc{ii}, newline));               %  删除工程名称字符串中的换行符，然后将文中'gcmc'替换为工程名称gcmc{ii}
                                                                            %   旧MatLab版本采用：“gcgk3 = strrep(gcgk2, 'gcmc' , strrep(gcmc{ii},char(10),'')); ”
    gcgk4 = strrep(gcgk3, 'jgwmc' , jgwmc{ii} );                            %  将文中'jgwmc'替换为结构物名称jgcmc{ii}
    selection.Text = gcgk4;
    selection.Style = '正文';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  回车，插入一行
    pilepositionfilespec = fullfile(pwd,'\报告中引用文件\',zwpmbzt{ii});     %   引用桩位平面布置图的路径及名称
    handle1 = selection.InlineShape.AddPicture(pilepositionfilespec);
    if picsize_1(3) == 1                                                    %  锁定图片宽高比缩放图片
        scaling = min( picsize_1(1)/handle1.Width , picsize_1(2)/handle1.Height );  %  图1的缩放比例
        handle1.Width = handle1.Width * scaling;                            % 等比例缩放图1的宽、高
        handle1.Height = handle1.Height * scaling;
    elseif picsize_1(3) == 0                                                %   不锁定图片宽高比缩放图片
        handle1.Width = picsize_1(1);
        handle1.Height = picsize_1(2);
    else
        error('picsize_1图片锁定宽高比参数输入有误(0或1)，请修改后重试！');
    end
    selection.Style = '表格';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  回车，插入一行
    selection.Text = '图 1  桩位平面布置图';
    selection.Style = '题注';
    selection.MoveDown;                                                    %  将光标移到所选区域的最后
    
    selection.TypeParagraph;                                                %  回车，插入一行
    selection.Text = '2.2地质条件描述';
    selection.Style = '标题 2';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  回车，插入一行
    selection.Text = dztj{ii};
    selection.Style = '正文';
    selection.MoveDown;
    
    selection.Start = content.end;                                          %  将选定区域的起始位置定位到文章末尾
    selection.InsertBreak(0);                                               %  插入分页符0或分节符2
    
    
    selection.Text = '3检测情况一览表';
    selection.Style = '标题 1';
    selection.MoveDown;
    
    selection.TypeParagraph;                                                %  回车，插入一行
    selection.Text = '表 2  检测情况一览表';
    selection.Style = '题注';
    selection.MoveDown;
    
    
    selection.TypeParagraph;
    document.Tables.Add(selection.Range,13,5);                              %  插入表格2
    numTables = numTables + 1;
    DTI = document.Tables.Item( numTables );
    
    for mm = 1:1:13
        DTI.Rows.Item(mm).Height = height_3(mm);                             %  设置表2单元格高度
        for nn = 1:1:5
            DTI.Columns.Item(nn).Width = width_3(nn);                        %  设置表2单元格宽度
            DTI.Cell(mm,nn).Range.Paragraphs.Alignment = 'wdAlignParagraphCenter';  %  设置单元格水平居中对齐
            DTI.Cell(mm,nn).VerticalAlignment = 'wdCellAlignVerticalCenter';        %  设置单元格对齐方式
            DTI.Cell(mm,nn).Range.ParagraphFormat.Style = '表格';            %   设置样式为‘表格’
        end
    end
    for mm = 1:2:7                                                         %  字体加黑
        DTI.Cell(mm,1).Range.Font.Bold = 1;
        DTI.Cell(mm,2).Range.Font.Bold = 1;
        DTI.Cell(mm,3).Range.Font.Bold = 1;
        DTI.Cell(mm,4).Range.Font.Bold = 1;
        DTI.Cell(mm,5).Range.Font.Bold = 1;
    end
    DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';                     %  设置外边框线型，查询用DTI.Borders.set('OutsideLineStyle')
    DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';                      %  设置外边框线宽
    DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';
    DTI.Borders.InsideLineWidth = 'wdLineWidth050pt';
    DTI.Rows.Alignment = 'wdAlignRowCenter';                                %  设
    DTI.Cell(8,1).Merge(DTI.Cell(13,1));                                    %  合并单元格，第1行2～4合并
    DTI.Cell(8,2).Merge(DTI.Cell(13,2));
    DTI.Cell(8,3).Merge(DTI.Cell(13,3));
    DTI.Cell(7,4).Merge(DTI.Cell(7,5));
    DTI.Cell(1,1).Range.Text = '工程名称';                                  %  录入表格文字
    DTI.Cell(2,1).Range.Text = gcmc{ii};
    DTI.Cell(1,2).Range.Text = '结构物名称';
    DTI.Cell(2,2).Range.Text = jgwmc{ii};
    DTI.Cell(1,3).Range.Text = '仪器设备';
    DTI.Cell(2,3).Range.Text = yqsb{ii};
    DTI.Cell(1,4).Range.Text = '检测标准';
    DTI.Cell(2,4).Range.Text = jcbz{ii};
    DTI.Cell(1,5).Range.Text = '基桩编号';
    DTI.Cell(2,5).Range.Text = jzbh{ii};    
    DTI.Cell(3,1).Range.Text = '设计桩顶标高(m)';                                  %  录入表格文字
    DTI.Cell(4,1).Range.Text = sjzTbg{ii};
    DTI.Cell(3,2).Range.Text = '设计桩端标高(m)';
    DTI.Cell(4,2).Range.Text = sjzBbg{ii};
    DTI.Cell(3,3).Range.Text = '设计桩长(m)';
    DTI.Cell(4,3).Range.Text = sjzc{ii};
    DTI.Cell(3,4).Range.Text = '设计桩径(mm)';                                %  录入表格文字
    DTI.Cell(4,4).Range.Text = sjzj{ii};
    DTI.Cell(3,5).Range.Text = '混凝土强度等级';
    DTI.Cell(4,5).Range.Text = concreteStrength{ii};
    DTI.Cell(5,1).Range.Text = '灌注日期';
    DTI.Cell(6,1).Range.Text = datestr( datenum( strtok(gzrq{ii}) ), 'yyyy-mm-dd' );
    DTI.Cell(5,2).Range.Text = '检测日期';
    DTI.Cell(6,2).Range.Text = datestr( datenum( strtok(jcrq{ii}) ), 'yyyy-mm-dd' );
    DTI.Cell(5,3).Range.Text = '实测桩长(m)';
    DTI.Cell(6,3).Range.Text = sczc{ii};
    DTI.Cell(5,4).Range.Text = '实测桩顶标高(m)';
    DTI.Cell(6,4).Range.Text = sczTbg{ii};
    
    DTI.Cell(5,5).Select();
    selection.Text = '灌注混凝土方量(m';
    selection.EndKey;
    selection.Text = '3';
    selection.Font.Superscript = 1;
    selection.EndKey;
    selection.Text = ')';
    selection.Font.Superscript = 0;
    
    DTI.Cell(6,5).Range.Text = gzfl{ii};
    DTI.Cell(7,1).Range.Text = '声测管方位示意图';
    DTI.Cell(7,2).Range.Text = '检测结果';
    DTI.Cell(8,2).Range.Text = jcjg{ii}';
    DTI.Cell(7,3).Range.Text = '缺陷情况分析';
    DTI.Cell(8,3).Range.Text = qxfx{ii}';
    DTI.Cell(7,4).Range.Text = '测管间距(mm)';
    DTI.Cell(8,4).Range.Text = '1-2';
    DTI.Cell(8,5).Range.Text = cgjj1_2{ii};
    DTI.Cell(9,4).Range.Text = '1-3';
    DTI.Cell(9,5).Range.Text = cgjj1_3{ii};
    DTI.Cell(10,4).Range.Text = '2-3';
    DTI.Cell(10,5).Range.Text = cgjj2_3{ii};
    
    if str2double(cggs{ii}) == 3                                         %  三根管 声测管方位示意图
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
    elseif str2double(cggs{ii}) == 4                                     %  四根管 声测管方位示意图
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
    
    selection.Text = strcat('附件：实测曲线、图表（共', fjys{ii}, '页）');
    selection.Style = '正文';
    selection.MoveDown;
    
    
end



%%%%%%%%%    设置页眉    %%%%%%%%%%
for mm = 1 : 2 : document.Sections.Count                                    %  中间数2为每2段执行一次循环
    disp(['    设置序号 ' , bgxh{beginNo+(mm-1)/2} , ' 报告编号( ', bgbh{beginNo+(mm-1)/2} ,' )页眉']);
    document.Sections.Item(mm).Headers.Item('wdHeaderFooterPrimary').LinkToPrevious = 0;                                                                 %  设置本节不链接到前一节页眉
    document.Sections.Item(mm).Headers.Item('wdHeaderFooterPrimary').Range.Text='';                                                                      %  设置本节页眉内容为''
    document.Sections.Item(mm).Headers.Item('wdHeaderFooterPrimary').Range.ParagraphFormat.Borders.Item('wdBorderBottom').LineStyle = 'wdLineStyleNone'; %  设置本节段落下划线为无
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').LinkToPrevious = 0;                                                               %  设置第二节不链接到前一节页眉
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.Text = strcat(jcnr{beginNo+(mm-1)/2},'报告',09,bgbh{beginNo+(mm-1)/2});   %   strcat('混凝土基桩完整性超声波检测',09,bgbh{beginNo+(mm-1)/2});         %   设置第二节页眉内容
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.Style = '页眉';                                                            %   设置第二节页眉样式为'页眉1'
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.ParagraphFormat.Borders.Item('wdBorderBottom').LineStyle = 'wdLineStyleThinThickSmallGap';   %  设置第二节段落下划线
    document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.ParagraphFormat.Borders.Item('wdBorderBottom').LineWidth = 'wdLineWidth150pt';               %  设置第二节段落下划线宽度
    tabstop = document.Sections.Item(mm+1).Headers.Item('wdHeaderFooterPrimary').Range.ParagraphFormat.TabStops.Add( headerFooterTabstop );                               %  设置第二段制表符位置
    tabstop.Alignment = 'wdAlignTabRight';                                                                                                                                %  设置制表符的对齐方式
end

%%%%%%%%    设置页脚    %%%%%%%%%%
for nn = 1 : 2 : document.Sections.Count
    disp(['    设置序号 ' , bgxh{beginNo+(nn-1)/2} , ' 报告编号( ', bgbh{beginNo+(nn-1)/2} ,' )页脚']);
    document.Sections.Item(nn).Footers.Item('wdHeaderFooterPrimary').LinkToPrevious = 0;            %  设置本节页脚不链接上一节页脚
    document.Sections.Item(nn).Footers.Item('wdHeaderFooterPrimary').Range.Text='';                 %  设置本节页脚内容
    document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').LinkToPrevious = 0;          %  设置第二节页脚不链接上一节页脚
    document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').Range.Fields.Add(document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').Range,[],'Page');   %  插入第二节页脚页码
    document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').Range.Select;
    selection.Range.InsertBefore( '第');
    selection.Range.InsertAfter(strcat( '页 共',bgys{beginNo+(nn-1)/2}, '页'));
    document.Sections.Item(nn+1).Footers.Item('wdHeaderFooterPrimary').Range.Style = '页脚';
    
    document.Sections.Item(nn).Footers.Item('wdHeaderFooterPrimary').PageNumbers.RestartNumberingAtSection = 1;  %  设置本节页码重开始计算
    document.Sections.Item(nn).Footers.Item('wdHeaderFooterPrimary').PageNumbers.StartingNumber = 1;             %  设置本节页码重开始计算为1
end
word.ActiveWindow.View.Type='wdPrintView';
document.ActiveWindow.ActivePane.View.SeekView ='wdSeekMainDocument';
word.ActiveWindow.View.Type='wdPrintView';


%%%%%%  保存文件  %%%%%%
disp('    保存报告页文档');
bgmc = strcat( '报告页', bgbh{beginNo}, '～' , bgbh{endNo}  );                         %  输出检测报告名称
filespec= fullfile( pwd , '\检测报告\' , bgmc );                              %  输出报告文件存放路径
try
    document.SaveAs(filespec);
catch
    document.SaveAs2(filespec);
end

if exportAsPDF == 1
    if  word.Version >= 12
        disp('    输出PDF格式文件');
        document.ExportAsFixedFormat( filespec , 17 );                            %  另存为PDF格式   17 - PDF格式，18 - XPS格式
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
%%   readParameters   读取excel文件中参数

raw1 = {};
raw2 = {};

for ii = 1 : 1 : size( raw1A , 1 )                                          %  去除Excel中空白行
    if ~isnan( raw1A{ii,1})                                                 %  判断第ii行第1个数字非空
        raw1 = [ raw1 ; raw1A(ii,:) ];
    end
end

for ii = 1 : 1 : size( raw2A , 1 )                                          %  去除Excel中空白行
    if ~isnan( raw2A{ii,1})                                                 %  判断第ii行第1个数字非空
        raw2 = [ raw2 ; raw2A(ii,:) ];
    end
end

[raw1_m, raw1_n] = size(raw1);
[raw2_m, raw2_n] = size(raw2);

for ii = 1:1:raw1_m
    for jj =1:1:raw1_n
        if isnan(raw1{ ii , jj } )
            raw1{ ii , jj } = char(32);                                    %  如果文件中没有填该项信息，自动补充为空格
        end
    end
end
for ii = 1:1:raw2_m
    for jj =1:1:raw2_n
        if isnan(raw2{ ii , jj } )
            raw2{ ii , jj } = char(32);                                    %  如果文件中没有填该项信息，自动补充为空格
        end
    end
end


bgxh = cell(raw2_m - 1, 1);      %  报告序号
bgbh = cell(raw2_m - 1, 1);      %  报告编号
jgwmc = cell(raw2_m - 1, 1);     %  结构物名称(桥名)
dtbh = cell(raw2_m - 1, 1);      %  墩台编号
jzbh = cell(raw2_m - 1, 1);      %  基桩编号(桩号)
concreteStrength = cell(raw2_m - 1, 1);      %  混凝土强度等级
sjzc = cell(raw2_m - 1, 1);      %  设计桩长(m)
sczc = cell(raw2_m - 1, 1);      %  实测桩长(m)
sjzj = cell(raw2_m - 1, 1);      %  设计桩径(mm)
sjzTbg = cell(raw2_m - 1, 1);    %  设计桩顶标高(m)
sjzBbg = cell(raw2_m - 1, 1);    %  设计桩端标高(m)
sczTbg = cell(raw2_m - 1, 1);    %  实测桩顶标高(m)
sczBbg = cell(raw2_m - 1, 1);    %  实测桩端标高(m)
gzfl = cell(raw2_m - 1, 1);      %  灌注方量(m3)
gzrq = cell(raw2_m - 1, 1);      %  灌注日期
jcrq = cell(raw2_m - 1, 1);      %  检测日期
jcjg = cell(raw2_m - 1, 1);      %  检测结果
qxfx = cell(raw2_m - 1, 1);      %  缺陷情况分析
cggs = cell(raw2_m - 1, 1);      %  测管根数
cgjj1_2 = cell(raw2_m - 1, 1);   %  测管间距1-2(mm)
cgjj1_3 = cell(raw2_m - 1, 1);   %  测管间距1-3(mm)
cgjj2_3 = cell(raw2_m - 1, 1);   %  测管间距2-3(mm)
cgjj1_4 = cell(raw2_m - 1, 1);   %  测管间距1-4(mm)
cgjj2_4 = cell(raw2_m - 1, 1);   %  测管间距2-4(mm)
cgjj3_4 = cell(raw2_m - 1, 1);   %  测管间距3-4(mm)
fjys = cell(raw2_m - 1, 1);      %  附件页数
cdjj = cell(raw2_m - 1, 1);      %  测点间距(m)
yqsb = cell(raw2_m - 1, 1);      %  仪器设备
jcbz = cell(raw2_m - 1, 1);      %  检测标准
jzbg = cell(raw2_m - 1, 1);      %  基准标高
gcmc = cell(raw2_m - 1, 1);      %  工程名称
wtdw = cell(raw2_m - 1, 1);      %  委托单位
jcnr = cell(raw2_m - 1, 1);      %  检测内容
jclb = cell(raw2_m - 1, 1);      %  检测类别
bxry = cell(raw2_m - 1, 1);      %  编写人员
jcry = cell(raw2_m - 1, 1);      %  检测人员
fhry = cell(raw2_m - 1, 1);      %  复核(签名)
shry = cell(raw2_m - 1, 1);      %  审核(签名)
pzry = cell(raw2_m - 1, 1);      %  批准(签名)
pzjb = cell(raw2_m - 1, 1);      %  批准级别
jcdw = cell(raw2_m - 1, 1);      %  检测单位
dwdz = cell(raw2_m - 1, 1);      %  地址
dwdh = cell(raw2_m - 1, 1);      %  电话
yzbm = cell(raw2_m - 1, 1);      %  邮政编码
dwcz = cell(raw2_m - 1, 1);      %  传真
bgrq = cell(raw2_m - 1, 1);      %  报告日期
bgys = cell(raw2_m - 1, 1);      %  报告页数
jsdw = cell(raw2_m - 1, 1);      %  建设单位
kcdw = cell(raw2_m - 1, 1);      %  勘察单位
sjdw = cell(raw2_m - 1, 1);      %  设计单位
zjb = cell(raw2_m - 1, 1);       %  总监办
jldw = cell(raw2_m - 1, 1);      %  监理单位
zdjl = cell(raw2_m - 1, 1);      %  驻地监理
sgdw = cell(raw2_m - 1, 1);      %  施工单位
jcxs = cell(raw2_m - 1, 1);      %  基础型式
sbjgxs = cell(raw2_m - 1, 1);    %  上部结构型式
jcxm = cell(raw2_m - 1, 1);      %  检测项目
jcff = cell(raw2_m - 1, 1);      %  检测方法
jcsl = cell(raw2_m - 1, 1);      %  检测数量
gcgk = cell(raw2_m - 1, 1);      %  工程概况
dztj = cell(raw2_m - 1, 1);      %  地质条件
sggc = cell(raw2_m - 1, 1);      %  施工过程
jcgc = cell(raw2_m - 1, 1);      %  检测过程
zwpmbzt = cell(raw2_m - 1, 1);   %  桩位平面布置图
bz = cell(raw2_m - 1, 1);        %  备注
zysx = cell(raw2_m - 1, 1);      %  注意事项

for ii = 1:1:raw1_m                                                        %  基本信息.xlsx 内容分类
    switch raw1{ii,1}
        case '序号'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bgxh{jj} = num2str(raw1{ii , 2});
                else
                    bgxh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '报告编号'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bgbh{jj} = num2str(raw1{ii , 2});
                else
                    bgbh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '结构物名称'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jgwmc{jj} = num2str(raw1{ii , 2});
                else
                    jgwmc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '墩台编号'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dtbh{jj} = num2str(raw1{ii , 2});
                else
                    dtbh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '基桩编号'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jzbh{jj} = num2str(raw1{ii , 2});
                else
                    jzbh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '混凝土强度等级'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    concreteStrength{jj} = num2str(raw1{ii , 2});
                else
                    concreteStrength{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '设计桩长(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjzc{jj} = num2str(raw1{ii , 2});
                else
                    sjzc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '实测桩长(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sczc{jj} = num2str(raw1{ii , 2});
                else
                    sczc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '设计桩径(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjzj{jj} = num2str(raw1{ii , 2});
                else
                    sjzj{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '设计桩顶标高(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjzTbg{jj} = num2str(raw1{ii , 2});
                else
                    sjzTbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '设计桩端标高(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjzBbg{jj} = num2str(raw1{ii , 2});
                else
                    sjzBbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '实测桩顶标高(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sczTbg{jj} = num2str(raw1{ii , 2});
                else
                    sczTbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '实测桩端标高(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sczBbg{jj} = num2str(raw1{ii , 2});
                else
                    sczBbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '灌注混凝土方量(m3)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    gzfl{jj} = num2str(raw1{ii , 2});
                else
                    gzfl{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '灌注日期'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    gzrq{jj} = num2str(raw1{ii , 2});
                else
                    gzrq{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测日期'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcrq{jj} = num2str(raw1{ii , 2});
                else
                    jcrq{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测结果'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcjg{jj} = num2str(raw1{ii , 2});
                else
                    jcjg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '缺陷情况分析'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    qxfx{jj} = num2str(raw1{ii , 2});
                else
                    qxfx{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '测管根数'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cggs{jj} = num2str(raw1{ii , 2});
                else
                    cggs{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '测管间距1-2(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj1_2{jj} = num2str(raw1{ii , 2});
                else
                    cgjj1_2{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '测管间距1-3(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj1_3{jj} = num2str(raw1{ii , 2});
                else
                    cgjj1_3{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '测管间距2-3(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj2_3{jj} = num2str(raw1{ii , 2});
                else
                    cgjj2_3{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '测管间距1-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj1_4{jj} = num2str(raw1{ii , 2});
                else
                    cgjj1_4{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '测管间距2-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj2_4{jj} = num2str(raw1{ii , 2});
                else
                    cgjj2_4{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '测管间距3-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cgjj3_4{jj} = num2str(raw1{ii , 2});
                else
                    cgjj3_4{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '附件页数'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    fjys{jj} = num2str(raw1{ii , 2});
                else
                    fjys{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '测点间距(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    cdjj{jj} = num2str(raw1{ii , 2});
                else
                    cdjj{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '仪器设备'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    yqsb{jj} = num2str(raw1{ii , 2});
                else
                    yqsb{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测标准'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcbz{jj} = num2str(raw1{ii , 2});
                else
                    jcbz{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '基准标高'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jzbg{jj} = num2str(raw1{ii , 2});
                else
                    jzbg{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '工程名称'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    gcmc{jj} = num2str(raw1{ii , 2});
                else
                    gcmc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '委托单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    wtdw{jj} = num2str(raw1{ii , 2});
                else
                    wtdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测内容'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcnr{jj} = num2str(raw1{ii , 2});
                else
                    jcnr{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测类别'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jclb{jj} = num2str(raw1{ii , 2});
                else
                    jclb{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '编写人员'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bxry{jj} = num2str(raw1{ii , 2});
                else
                    bxry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测人员'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcry{jj} = num2str(raw1{ii , 2});
                else
                    jcry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '复核(签名)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    fhry{jj} = num2str(raw1{ii , 2});
                else
                    fhry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '审核(签名)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    shry{jj} = num2str(raw1{ii , 2});
                else
                    shry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '批准(签名)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    pzry{jj} = num2str(raw1{ii , 2});
                else
                    pzry{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '批准级别'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    pzjb{jj} = num2str(raw1{ii , 2});
                else
                    pzjb{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcdw{jj} = num2str(raw1{ii , 2});
                else
                    jcdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '地址'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dwdz{jj} = num2str(raw1{ii , 2});
                else
                    dwdz{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '电话'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dwdh{jj} = num2str(raw1{ii , 2});
                else
                    dwdh{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '邮政编码'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    yzbm{jj} = num2str(raw1{ii , 2});
                else
                    yzbm{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '传真'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dwcz{jj} = num2str(raw1{ii , 2});
                else
                    dwcz{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '报告日期'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bgrq{jj} = num2str(raw1{ii , 2});
                else
                    bgrq{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '报告页数'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bgys{jj} = num2str(raw1{ii , 2});
                else
                    bgys{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '建设单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jsdw{jj} = num2str(raw1{ii , 2});
                else
                    jsdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '勘察单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    kcdw{jj} = num2str(raw1{ii , 2});
                else
                    kcdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '设计单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sjdw{jj} = num2str(raw1{ii , 2});
                else
                    sjdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '总监办'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    zjb{jj} = num2str(raw1{ii , 2});
                else
                    zjb{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '监理单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jldw{jj} = num2str(raw1{ii , 2});
                else
                    jldw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '驻地监理'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    zdjl{jj} = num2str(raw1{ii , 2});
                else
                    zdjl{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '施工单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sgdw{jj} = num2str(raw1{ii , 2});
                else
                    sgdw{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '基础型式'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcxs{jj} = num2str(raw1{ii , 2});
                else
                    jcxs{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '上部结构型式'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sbjgxs{jj} = num2str(raw1{ii , 2});
                else
                    sbjgxs{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测项目'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcxm{jj} = num2str(raw1{ii , 2});
                else
                    jcxm{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测方法'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcff{jj} = num2str(raw1{ii , 2});
                else
                    jcff{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测数量'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcsl{jj} = num2str(raw1{ii , 2});
                else
                    jcsl{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '工程概况'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    gcgk{jj} = num2str(raw1{ii , 2});
                else
                    gcgk{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '地质条件'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    dztj{jj} = num2str(raw1{ii , 2});
                else
                    dztj{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '施工过程'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    sggc{jj} = num2str(raw1{ii , 2});
                else
                    sggc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '检测过程'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    jcgc{jj} = num2str(raw1{ii , 2});
                else
                    jcgc{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '桩位平面布置图'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    zwpmbzt{jj} = num2str(raw1{ii , 2});
                else
                    zwpmbzt{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '备注'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    bz{jj} = num2str(raw1{ii , 2});
                else
                    bz{jj} = raw1{ii , 2};
                end
            end
            continue;
        case '注意事项'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw1{ii , 2})
                    zysx{jj} = num2str(raw1{ii , 2});
                else
                    zysx{jj} = raw1{ii , 2};
                end
            end
            continue;
            
        otherwise
            disp('未利用参数：');
            disp(['    ',raw1{ii,1}]);
    end
end


for ii = 1:1:raw2_n                                                        %  检测信息.xlsx 内容分类
    switch raw2{1,ii}
        case '序号'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bgxh{jj} = num2str(raw2{jj+1 , ii});
                else
                    bgxh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '报告编号'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bgbh{jj} = num2str(raw2{jj+1 , ii});
                else
                    bgbh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '结构物名称'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jgwmc{jj} = num2str(raw2{jj+1 , ii});
                else
                    jgwmc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '墩台编号'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dtbh{jj} = num2str(raw2{jj+1 , ii});
                else
                    dtbh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '基桩编号'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jzbh{jj} = num2str(raw2{jj+1 , ii});
                else
                    jzbh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '混凝土强度等级'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    concreteStrength{jj} = num2str(raw2{jj+1 , ii});
                else
                    concreteStrength{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '设计桩长(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjzc{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjzc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '实测桩长(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sczc{jj} = num2str(raw2{jj+1 , ii});
                else
                    sczc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '设计桩径(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjzj{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjzj{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '设计桩顶标高(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjzTbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjzTbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '设计桩端标高(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjzBbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjzBbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '实测桩顶标高(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sczTbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    sczTbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '实测桩端标高(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sczBbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    sczBbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;   
        case '灌注混凝土方量(m3)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    gzfl{jj} = num2str(raw2{jj+1 , ii});
                else
                    gzfl{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '灌注日期'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    gzrq{jj} = num2str(raw2{jj+1 , ii});
                else
                    gzrq{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测日期'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcrq{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcrq{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测结果'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcjg{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcjg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '缺陷情况分析'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    qxfx{jj} = num2str(raw2{jj+1 , ii});
                else
                    qxfx{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '测管根数'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cggs{jj} = num2str(raw2{jj+1 , ii});
                else
                    cggs{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '测管间距1-2(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj1_2{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj1_2{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '测管间距1-3(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj1_3{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj1_3{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '测管间距2-3(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj2_3{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj2_3{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '测管间距1-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj1_4{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj1_4{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '测管间距2-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj2_4{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj2_4{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '测管间距3-4(mm)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cgjj3_4{jj} = num2str(raw2{jj+1 , ii});
                else
                    cgjj3_4{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '附件页数'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    fjys{jj} = num2str(raw2{jj+1 , ii});
                else
                    fjys{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '测点间距(m)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    cdjj{jj} = num2str(raw2{jj+1 , ii});
                else
                    cdjj{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '仪器设备'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    yqsb{jj} = num2str(raw2{jj+1 , ii});
                else
                    yqsb{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测标准'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcbz{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcbz{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '基准标高'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jzbg{jj} = num2str(raw2{jj+1 , ii});
                else
                    jzbg{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '工程名称'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    gcmc{jj} = num2str(raw2{jj+1 , ii});
                else
                    gcmc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '委托单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    wtdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    wtdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测内容'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcnr{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcnr{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测类别'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jclb{jj} = num2str(raw2{jj+1 , ii});
                else
                    jclb{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '编写人员'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bxry{jj} = num2str(raw2{jj+1 , ii});
                else
                    bxry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测人员'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcry{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '复核(签名)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    fhry{jj} = num2str(raw2{jj+1 , ii});
                else
                    fhry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '审核(签名)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    shry{jj} = num2str(raw2{jj+1 , ii});
                else
                    shry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '批准(签名)'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    pzry{jj} = num2str(raw2{jj+1 , ii});
                else
                    pzry{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '批准级别'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    pzjb{jj} = num2str(raw2{jj+1 , ii});
                else
                    pzjb{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '地址'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dwdz{jj} = num2str(raw2{jj+1 , ii});
                else
                    dwdz{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '电话'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dwdh{jj} = num2str(raw2{jj+1 , ii});
                else
                    dwdh{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '邮政编码'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    yzbm{jj} = num2str(raw2{jj+1 , ii});
                else
                    yzbm{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '传真'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dwcz{jj} = num2str(raw2{jj+1 , ii});
                else
                    dwcz{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '报告日期'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bgrq{jj} = num2str(raw2{jj+1 , ii});
                else
                    bgrq{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '报告页数'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bgys{jj} = num2str(raw2{jj+1 , ii});
                else
                    bgys{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '建设单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jsdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    jsdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '勘察单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    kcdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    kcdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '设计单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sjdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    sjdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '总监办'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    zjb{jj} = num2str(raw2{jj+1 , ii});
                else
                    zjb{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '监理单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jldw{jj} = num2str(raw2{jj+1 , ii});
                else
                    jldw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '驻地监理'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    zdjl{jj} = num2str(raw2{jj+1 , ii});
                else
                    zdjl{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '施工单位'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sgdw{jj} = num2str(raw2{jj+1 , ii});
                else
                    sgdw{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '基础型式'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcxs{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcxs{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '上部结构型式'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sbjgxs{jj} = num2str(raw2{jj+1 , ii});
                else
                    sbjgxs{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测项目'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcxm{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcxm{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测方法'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcff{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcff{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测数量'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcsl{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcsl{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '工程概况'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    gcgk{jj} = num2str(raw2{jj+1 , ii});
                else
                    gcgk{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '地质条件'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    dztj{jj} = num2str(raw2{jj+1 , ii});
                else
                    dztj{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '施工过程'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    sggc{jj} = num2str(raw2{jj+1 , ii});
                else
                    sggc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '检测过程'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    jcgc{jj} = num2str(raw2{jj+1 , ii});
                else
                    jcgc{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '桩位平面布置图'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    zwpmbzt{jj} = num2str(raw2{jj+1 , ii});
                else
                    zwpmbzt{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '备注'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    bz{jj} = num2str(raw2{jj+1 , ii});
                else
                    bz{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
        case '注意事项'
            for jj = 1 : 1 : raw2_m - 1
                if isnumeric(raw2{jj+1 , ii})
                    zysx{jj} = num2str(raw2{jj+1 , ii});
                else
                    zysx{jj} = raw2{jj+1 , ii};
                end
            end
            continue;
            
        otherwise
            disp('未利用参数：');
            disp(['    ',raw2{1,ii}]);
    end
end


end