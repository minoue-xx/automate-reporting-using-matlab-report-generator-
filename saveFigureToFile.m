% Copyright (c) 2023 Michio Inoue.

function imgPath = saveFigureToFile(hFigure)

% Create images folder if it does not exist
imgFolderPath = fullfile('.','tmp_plots');
if( ~isfolder(imgFolderPath) )
    mkdir(imgFolderPath);
end

% Create randomized name for image file
imgName = ['img_',char(datetime('now','Format','yyyyMMddHHmmssSSSSS'))];

% Select an appropriate image type depending on the platform.
if ~ispc
    imgType = '-dpng';
    imgName= [imgName '.png'];
else
    % This Microsoft-specific vector graphics format
    % can yield better quality images in Word documents.
    imgType = '-dmeta';
    imgName = [imgName '.emf'];
end

% Create the path for the file
imgPath = fullfile(imgFolderPath,imgName);

% Save image to file
print(hFigure,imgPath,imgType);

end
