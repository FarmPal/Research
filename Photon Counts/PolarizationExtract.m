% PolarizationExtract
% Grab and sum the central intensities from each sheet in an excel sheet
% Created on Mar. 30, 2015 by Jason Ng

excelFileName = 'Data.xls';
allData = loadExcel(excelFileName);    % Grab the data from the proper excel file 
numPics = size(allData,3);              % Number of data points
intensities = zeros(numPics,1);      % Holds output sums

% Set the input ranges
rangeX1 = floor(size(allData,1)*(1/4));
rangeX2 = ceil(size(allData,1)*(3/4));
rangeY1 = floor(size(allData,2)*(1/4));
rangeY2 = ceil(size(allData,2)*(3/4));

for i = 1:1:numPics
    intensities(i) = sum(sum(allData(rangeX1:rangeX2,rangeY1:rangeY2,i)));
end

xlswrite('Polarization.xls',intensities,'Intensities');

% Plot the range over the first image
plotData = allData(1:255,1:255,n);
pcolor(plotData); % Omit the last column
rectangle('Position',[rangeX1 rangeY1 rangeX2-rangeX1 rangeY2-rangeY1]);

% plot(intensities);

% Experiment with different 3D plotting options
% mesh / surf - draws with 3D perspective, with diff colours/lines
% pcolor - gives 2D image with colours
%   - Images are right side up, without axis markers
%   - Seems to be better than image/imagesc

shading flat;   % Removes grid lines and looks sharper
colormap(bone);