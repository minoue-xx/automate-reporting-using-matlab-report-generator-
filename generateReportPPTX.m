% Copyright (c) 2023 Michio Inoue.

%% 事前設定
% ファイル名に今日の日付を使います
reportDate = datetime; 
reportDate.Format = 'yyyy-MM-dd';

% 関連ライブラリの読み込み
import mlreportgen.ppt.*;

% 事前作成したテンプレートからスライド作成
pres = Presentation(fileName,'sampleTemplate.potx');

%% タイトルスライド作成
% テンプレートにある "Title Slide" レイアウトを使用
titleSlide = add(pres,'Title Slide');

% ファイル名やタイトル
fileName     = "sampleReport" + string(reportDate);
titleText    = "サンプルレポート";
subTitleText = "作成日" + string(reportDate);

% プレースフォルダの内容を入れ替え
replace(titleSlide,'Title',titleText);
replace(titleSlide,'Subtitle',subTitleText);

%% ここからレポートページ
% さいころを N 回振って、出る目の回数を集計します。
names = ["one","two","three","four","five","six"];

for ii=1:5 % 10回 から 10万回まで
    N = 10^ii;
    
    % 1 から 6 までの整数を生成
    diceResults = randi(6,[N,1]);
    % それぞれの出目の回数を集計
    counts = histcounts(diceResults);
    [countsSorted,idx] = sort(counts,'descend');
    
    % ヒストグラムプロット作成
    hFigure = figure(1);
    histogram(diceResults,'Normalization','probability');
    title('Histogram of Each Occurenace');
    ha = gca;
    ha.FontSize = 20;
    
    % 画像として保存して後に、パワポにコピー
    imgPath = saveFigureToFile(hFigure);
    pictureObj = Picture(imgPath);
    
    % スライドタイトル
    slideTitle = "Uniformly Distributed? (N = " + string(N) + ")";
    
    % 事前に準備しておいた Custom レイアウトからページ作成
    slide = add(pres, 'Custom');
    replace(slide,'Title',slideTitle); % タイトル
    replace(slide,'Picture Placeholder Big', pictureObj); % 真ん中はヒストグラムを配置
    
    % 出た回数順にトップ５：出た回数と画像を配置
    topfive = names(idx);
    for jj=1:5
        % 画像は事前に用意したものを使用
        imagePath = "./images/" + topfive(jj) + ".png";
        if ~exist(imagePath,'file') % 念のため画像がない場合
            imagePath = "./images/a.png";
        end
        pictureObj = Picture(char(imagePath));
        replace(slide,"Picture Placeholder " + jj, pictureObj);
        replace(slide,"Text Placeholder " + jj, string(countsSorted(ii)));
    end
    
    % hFigure を閉じる
    if( isvalid(hFigure) )
        close(hFigure);
    end
end

%% 以上のページをもとにパワポ作成
% msgbox でダイアログ出してみる
hMsg = msgbox('Generating PowerPoint report...');

% pres を閉じておきます。
close(pres);
if(ispc) % できたパワポを開く
    winopen(pres.OutputPath);
end

% ダイアログ閉じる
if( isvalid(hMsg) )
    close(hMsg);
end
