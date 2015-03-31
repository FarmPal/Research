% PCExtract
% Photon counting extraction data to convert SPC data into usable matrices
% Created on Mar. 23, 2015 by Jason Ng

file = importdata ('filename.txt');
numFiles = size(file, 1);
rawData = {};                   % Raw data input
inData = zeros(65536,256);      % Specify expected size of output
pixelSum = zeros(65536,1);
output2D = zeros(256,256);
outputClipped = zeros(256,255);
excelFileName = 'Data.xls';

allData = zeros(256,256,numFiles);  % Hold all the data
tabNames = {};                      % Hold all the tab names
    % Double check whether the 2nd and 3rd arguments are necessary

% Specific data for importing photon count data
delimiter = ' ';        % Delimiter character
numHeaderLines = 10;     % # of header lines in photon count data files
dataPerRow = 256;       % # of data per pixel


% For each file, create a separate excel file within the loop
% 256 data units per row first
for n = 1 : 1 : numFiles %loop for searching the file
%for n = 1 : 1 : 1 %loop for searching the file
    filename = char(file(n));
    
    rawData = importdata(filename, delimiter, numHeaderLines);
    inData = vec2mat (rawData.data,256);  % Converts data into 256^2x256
    pixelSum = sum(inData,2);
    output2D = vec2mat(pixelSum,256);
    outputClipped = output2D(:,1:255);
    
    % Save all the data into a super matrix
    allData(:,:,n) = output2D;    
        
    % Write the data into 
    %newName = [filename(1:(size(filename,2)-4)) '.xls'];
    tabName = [filename(1:3) filename(49:size(filename,2))];
    tabName = tabName(1:size(tabName,2)-4);
    tabNames = [tabNames tabName];

    %xlswrite('Data.xls', output2D, tabName); 
end

% Create the Excel file with all the data and tab names 
for n = 1:1:numFiles
    xlswrite(excelFileName, allData(:,:,n), char(tabNames(1,n)));
end

% Create png's for each picture
for n = 1:1:numFiles
    plotData = allData(1:255,1:255,n);
    plotData(1,1) = 120;            % Maintain consistent scaling (120 for non-IR)
    pcolor(plotData); % Omit the last column
    colormap(bone);
    shading flat;
    caxis([2,100]);
    print([char(tabNames(1,n)) '.png'],'-dpng');
end

    
% Remove the original blank sheets
% http://www.mathworks.com/matlabcentral/answers/92449-how-can-i-delete-the
% -default-sheets-sheet1-sheet2-and-sheet3-in-excel-when-i-use-xlswrite
% excelFileName = 'Data.xls'; // Defined up top
excelFilePath = pwd; % Current working directory.
sheetName = 'Sheet'; % EN: Sheet, DE: Tabelle, etc. (Lang. dependent)
% Open Excel file.
objExcel = actxserver('Excel.Application');
objExcel.Workbooks.Open(fullfile(excelFilePath, excelFileName)); % Full path is necessary!
% Delete sheets.
try
      % Throws an error if the sheets do not exist.
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '1']).Delete;
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '2']).Delete;
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '3']).Delete;
catch
      ; % Do nothing.
end
% Save, close and clean up.
objExcel.ActiveWorkbook.Save;
objExcel.ActiveWorkbook.Close;
objExcel.Quit;
objExcel.delete;
