clear
%formatted for data in "Summary" files

files = dir('C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\summary\*.xls*');
cd 'C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\summary';
frontal_info = {length(files), 12};

for n = 1:length(files)
    f = files(n).name; %'CEF1007 Summary.xlsm';

    %% Test & vehicle data
    testInfo = readtable(f, 'Sheet', 2, 'Range', 'A2:B2', 'ReadVariableNames', false);
    testID = char(string(testInfo.Var1(1)));
    vehInfo = char(string(testInfo.Var2(1)));

    peakData = readtable(f, 'Sheet', 2, 'VariableNamingRule', 'preserve', 'ReadRowNames', true);
    vehRowNames = {'10VEHC0000__ACRD', 'VEHC0000__VEX'};
    rowNamesHead = {'11HEADCG00__ACXA', '11HEADCG00__ACYA', '11HEADCG00__ACZA', '11HEADCG00__ACR'};
    rowNamesChest = {'11CHST0000__ACXC', '11CHST0000__ACYC', '11CHST0000__ACZC', '11CHST0000__ACRC'};

    rowNames = [rowNamesHead, rowNamesChest];

    %peak vehicle a and dV (dV calculated based on spreadsheet)
    vehAPeak = max(abs(peakData{vehRowNames{1}, 4}), abs(peakData{vehRowNames{1}, 5}));
    vehdvPeak = peakData{vehRowNames{2}, 5} - peakData{vehRowNames{2}, 4};

    %% Head accel data
    headAXPeak = max(abs(peakData{rowNames{1}, 4}), abs(peakData{rowNames{1}, 5}));
    headAYPeak = max(abs(peakData{rowNames{2}, 4}), abs(peakData{rowNames{2}, 5}));
    headAZPeak = max(abs(peakData{rowNames{3}, 4}), abs(peakData{rowNames{3}, 5}));
    headARPeak = max(abs(peakData{rowNames{4}, 4}), abs(peakData{rowNames{4}, 5}));
    
    %% Chest accel data
    chestAXPeak = max(abs(peakData{rowNames{5}, 4}), abs(peakData{rowNames{5}, 5}));
    chestAYPeak = max(abs(peakData{rowNames{6}, 4}), abs(peakData{rowNames{6}, 5}));
    chestAZPeak = max(abs(peakData{rowNames{7}, 4}), abs(peakData{rowNames{7}, 5}));
    chestARPeak = max(abs(peakData{rowNames{8}, 4}), abs(peakData{rowNames{8}, 5}));

    %% add all data to cell array (for export to excel later)
    frontal_info{n, 1} = testID;
    frontal_info{n, 2} = vehInfo;
    frontal_info{n, 3} = vehdvPeak;
    frontal_info{n, 4} = vehAPeak;
    frontal_info{n, 5} = headAXPeak;
    frontal_info{n, 6} = headAYPeak;
    frontal_info{n, 7} = headAZPeak;
    frontal_info{n, 8} = headARPeak;
    frontal_info{n, 9} = chestAXPeak;
    frontal_info{n, 10} = chestAYPeak;
    frontal_info{n, 11} = chestAZPeak;
    frontal_info{n, 12} = chestARPeak;
end
save('IIHS_summary.mat', 'frontal_info');
close all
