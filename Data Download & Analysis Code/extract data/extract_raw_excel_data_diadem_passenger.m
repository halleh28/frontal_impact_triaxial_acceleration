clear
%formatted for data in "diadem\passenger" files

files = dir('C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\passenger\*.xls*');
cd 'C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\passenger';
frontal_info = {length(files), 10};

for n = 1:length(files)
    f = files(n).name; %'CF16005 (Passenger).xlsx';
    [~, sheetNames] = xlsfinfo(f);

    %% Test & vehicle data
    testInfo = readtable(f, 'Sheet', 1, 'Range', 'A2:B2', 'ReadVariableNames', false);
    testID = char(string(testInfo.Var1(1)));
    vehInfo = char(string(testInfo.Var2(1)));

    peakData = readtable(f, 'Sheet', 1, 'VariableNamingRule', 'preserve', 'ReadRowNames', true);
    pRowNamesHead = {'13HEADCG00__ACXA', '13HEADCG00__ACYA', '13HEADCG00__ACZA', '13HEADCG00__ACR'};
    pRowNamesChest = {'13CHST0000__ACXC', '13CHST0000__ACYC', '13CHST0000__ACZC', '13CHST0000__ACRC'};

    pRowNames = [pRowNamesHead, pRowNamesChest];

    %% Head accel data
    headAXPeak = max(abs(peakData{pRowNames{1}, 4}), abs(peakData{pRowNames{1}, 5}));
    headAYPeak = max(abs(peakData{pRowNames{2}, 4}), abs(peakData{pRowNames{2}, 5}));
    headAZPeak = max(abs(peakData{pRowNames{3}, 4}), abs(peakData{pRowNames{3}, 5}));
    headARPeak = max(abs(peakData{pRowNames{4}, 4}), abs(peakData{pRowNames{4}, 5}));

    %% Chest accel data
    chestAXPeak = max(abs(peakData{pRowNames{5}, 4}), abs(peakData{pRowNames{5}, 5}));
    chestAYPeak = max(abs(peakData{pRowNames{6}, 4}), abs(peakData{pRowNames{6}, 5}));
    chestAZPeak = max(abs(peakData{pRowNames{7}, 4}), abs(peakData{pRowNames{7}, 5}));
    chestARPeak = max(abs(peakData{pRowNames{8}, 4}), abs(peakData{pRowNames{8}, 5}));

    %% add all data to cell array (for export to excel later)
    frontal_info{n, 1} = testID;
    frontal_info{n, 2} = vehInfo;
    frontal_info{n, 3} = headAXPeak;
    frontal_info{n, 4} = headAYPeak;
    frontal_info{n, 5} = headAZPeak;
    frontal_info{n, 6} = headARPeak;
    frontal_info{n, 7} = chestAXPeak;
    frontal_info{n, 8} = chestAYPeak;
    frontal_info{n, 9} = chestAZPeak;
    frontal_info{n, 10} = chestARPeak;
end
save('IIHS_diadem_p.mat', 'frontal_info');
close all
