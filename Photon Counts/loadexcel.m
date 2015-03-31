% LoadExcel
% Load data from the excel sheet into the originally named matrices
% Created on Mar. 30, 2015 by Jason Ng

function allData = loadexcel(excelFileName)
% Figure out the number of sheets
%excelFileName = 'Data.xls';    % Grab this as an input
excelFilePath = pwd; % Current working directory.
objExcel = actxserver('Excel.Application');
efile = objExcel.Workbooks.Open(fullfile(excelFilePath, excelFileName)); % Full path is necessary!
numSheets = efile.Worksheets.Count;
efile.Close;

allData = zeros(256,256,numSheets);

% Extract all the sheet data into allData
for i = 1:1:numSheets
    allData(:,:,i) = xlsread(excelFileName,i);
end