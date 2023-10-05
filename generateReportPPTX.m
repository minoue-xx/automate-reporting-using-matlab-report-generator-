% Copyright (c) 2023 Michio Inoue.

%% ���O�ݒ�
% �t�@�C�����ɍ����̓��t���g���܂�
reportDate = datetime; 
reportDate.Format = 'yyyy-MM-dd';

% �֘A���C�u�����̓ǂݍ���
import mlreportgen.ppt.*;

% ���O�쐬�����e���v���[�g����X���C�h�쐬
pres = Presentation(fileName,'sampleTemplate.potx');

%% �^�C�g���X���C�h�쐬
% �e���v���[�g�ɂ��� "Title Slide" ���C�A�E�g���g�p
titleSlide = add(pres,'Title Slide');

% �t�@�C������^�C�g��
fileName     = "sampleReport" + string(reportDate);
titleText    = "�T���v�����|�[�g";
subTitleText = "�쐬��" + string(reportDate);

% �v���[�X�t�H���_�̓��e�����ւ�
replace(titleSlide,'Title',titleText);
replace(titleSlide,'Subtitle',subTitleText);

%% �������烌�|�[�g�y�[�W
% ��������� N ��U���āA�o��ڂ̉񐔂��W�v���܂��B
names = ["one","two","three","four","five","six"];

for ii=1:5 % 10�� ���� 10����܂�
    N = 10^ii;
    
    % 1 ���� 6 �܂ł̐����𐶐�
    diceResults = randi(6,[N,1]);
    % ���ꂼ��̏o�ڂ̉񐔂��W�v
    counts = histcounts(diceResults);
    [countsSorted,idx] = sort(counts,'descend');
    
    % �q�X�g�O�����v���b�g�쐬
    hFigure = figure(1);
    histogram(diceResults,'Normalization','probability');
    title('Histogram of Each Occurenace');
    ha = gca;
    ha.FontSize = 20;
    
    % �摜�Ƃ��ĕۑ����Č�ɁA�p���|�ɃR�s�[
    imgPath = saveFigureToFile(hFigure);
    pictureObj = Picture(imgPath);
    
    % �X���C�h�^�C�g��
    slideTitle = "Uniformly Distributed? (N = " + string(N) + ")";
    
    % ���O�ɏ������Ă����� Custom ���C�A�E�g����y�[�W�쐬
    slide = add(pres, 'Custom');
    replace(slide,'Title',slideTitle); % �^�C�g��
    replace(slide,'Picture Placeholder Big', pictureObj); % �^�񒆂̓q�X�g�O������z�u
    
    % �o���񐔏��Ƀg�b�v�T�F�o���񐔂Ɖ摜��z�u
    topfive = names(idx);
    for jj=1:5
        % �摜�͎��O�ɗp�ӂ������̂��g�p
        imagePath = "./images/" + topfive(jj) + ".png";
        if ~exist(imagePath,'file') % �O�̂��߉摜���Ȃ��ꍇ
            imagePath = "./images/a.png";
        end
        pictureObj = Picture(char(imagePath));
        replace(slide,"Picture Placeholder " + jj, pictureObj);
        replace(slide,"Text Placeholder " + jj, string(countsSorted(ii)));
    end
    
    % hFigure �����
    if( isvalid(hFigure) )
        close(hFigure);
    end
end

%% �ȏ�̃y�[�W�����ƂɃp���|�쐬
% msgbox �Ń_�C�A���O�o���Ă݂�
hMsg = msgbox('Generating PowerPoint report...');

% pres ����Ă����܂��B
close(pres);
if(ispc) % �ł����p���|���J��
    winopen(pres.OutputPath);
end

% �_�C�A���O����
if( isvalid(hMsg) )
    close(hMsg);
end
