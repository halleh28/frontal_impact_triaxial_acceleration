clear
%formatted for data in "diadem\driver + passenger" files

files = dir('C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\driver + passenger\*.xls*');
cd 'C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\driver + passenger';
frontal_info = {length(files), 24};

for n = 1:length(files)
    f = files(n).name; %'CEF2103.xlsx';
    [~, sheetNames] = xlsfinfo(f);

    %% Test & vehicle data
    testInfo = readtable(f, 'Sheet', 1, 'Range', 'A2:R2', 'ReadVariableNames', false);
    testID = char(string(testInfo.Var1(1)));
    vehInfo = char(string(testInfo.Var18(1)));

    peakData = readtable(f, 'Sheet', 1, 'VariableNamingRule', 'preserve', 'ReadRowNames', true);
    dRowNamesHead = {'11HEADCG00__ACXA', '11HEADCG00__ACYA', '11HEADCG00__ACZA', '11HEADCG00__ACR'};
    dRowNamesChest = {'11CHST0000__ACXC', '11CHST0000__ACYC', '11CHST0000__ACZC', '11CHST0000__ACRC'};
    pRowNamesHead = {'14HEAD0000HFACXA', '14HEAD0000HFACYA', '14HEAD0000HFACZA', '14HEAD0000HFACR'};
    pRowNamesChest = {'14CHST0000HFACXC', '14CHST0000HFACYC', '14CHST0000HFACZC', '14CHST0000HFACRC'};
    pRowNamesPelvis = {'14PELV0000HFACXA', '14PELV0000HFACYA', '14PELV0000HFACZA', '14PELV0000HFACR'};
    vehRowNames = {'10VEHC0000__ACRD', 'VEHC0000__VEX'};
    vehTrunkRowNames = {'10VEHCTRNK__ACRD', 'VEHCTRNK__VEX'};

    dRowNames = [dRowNamesHead, dRowNamesChest];
    pRowNames = [pRowNamesHead, pRowNamesChest, pRowNamesPelvis];

    %peak vehicle a and dV (dV calculated based on spreadsheet)
    if isnan(peakData{vehRowNames{1}, 5})
        vehAPeak = max(abs(peakData{vehTrunkRowNames{1}, 4}), abs(peakData{vehTrunkRowNames{1}, 5}));
        vehdvPeak = peakData{vehTrunkRowNames{2}, 5} - peakData{vehTrunkRowNames{2}, 4};
    else
        vehAPeak = max(abs(peakData{vehRowNames{1}, 4}), abs(peakData{vehRowNames{1}, 5}));
        vehdvPeak = peakData{vehRowNames{2}, 5} - peakData{vehRowNames{2}, 4};
    end

    %% Driver head accel data
    dHeadAXPeak = max(abs(peakData{dRowNames{1}, 4}), abs(peakData{dRowNames{1}, 5}));
    dHeadAYPeak = max(abs(peakData{dRowNames{2}, 4}), abs(peakData{dRowNames{2}, 5}));
    dHeadAZPeak = max(abs(peakData{dRowNames{3}, 4}), abs(peakData{dRowNames{3}, 5}));
    dHeadARPeak = max(abs(peakData{dRowNames{4}, 4}), abs(peakData{dRowNames{4}, 5}));

    %% Driver chest accel data
    dChestAXPeak = max(abs(peakData{dRowNames{5}, 4}), abs(peakData{dRowNames{5}, 5}));
    dChestAYPeak = max(abs(peakData{dRowNames{6}, 4}), abs(peakData{dRowNames{6}, 5}));
    dChestAZPeak = max(abs(peakData{dRowNames{7}, 4}), abs(peakData{dRowNames{7}, 5}));
    dChestARPeak = max(abs(peakData{dRowNames{8}, 4}), abs(peakData{dRowNames{8}, 5}));

    %% Passenger head accel data
    pHeadAXPeak = max(abs(peakData{pRowNames{1}, 4}), abs(peakData{pRowNames{1}, 5}));
    pHeadAYPeak = max(abs(peakData{pRowNames{2}, 4}), abs(peakData{pRowNames{2}, 5}));
    pHeadAZPeak = max(abs(peakData{pRowNames{3}, 4}), abs(peakData{pRowNames{3}, 5}));
    pHeadARPeak = max(abs(peakData{pRowNames{4}, 4}), abs(peakData{pRowNames{4}, 5}));

    %% Passenger chest accel data
    pChestAXPeak = max(abs(peakData{pRowNames{5}, 4}), abs(peakData{pRowNames{5}, 5}));
    pChestAYPeak = max(abs(peakData{pRowNames{6}, 4}), abs(peakData{pRowNames{6}, 5}));
    pChestAZPeak = max(abs(peakData{pRowNames{7}, 4}), abs(peakData{pRowNames{7}, 5}));
    pChestARPeak = max(abs(peakData{pRowNames{8}, 4}), abs(peakData{pRowNames{8}, 5}));  

    %% Passenger pelvis accel data
    pPelvisAXPeak = max(abs(peakData{pRowNames{9}, 4}), abs(peakData{pRowNames{9}, 5}));
    pPelvisAYPeak = max(abs(peakData{pRowNames{10}, 4}), abs(peakData{pRowNames{10}, 5}));
    pPelvisAZPeak = max(abs(peakData{pRowNames{11}, 4}), abs(peakData{pRowNames{11}, 5}));
    pPelvisARPeak = max(abs(peakData{pRowNames{12}, 4}), abs(peakData{pRowNames{12}, 5}));

    %% add all data to cell array (for export to excel later)
    frontal_info{n, 1} = testID;
    frontal_info{n, 2} = vehInfo;
    frontal_info{n, 3} = vehdvPeak;
    frontal_info{n, 4} = vehAPeak;
    frontal_info{n, 5} = dHeadAXPeak;
    frontal_info{n, 6} = dHeadAYPeak;
    frontal_info{n, 7} = dHeadAZPeak;
    frontal_info{n, 8} = dHeadARPeak;
    frontal_info{n, 9} = dChestAXPeak;
    frontal_info{n, 10} = dChestAYPeak;
    frontal_info{n, 11} = dChestAZPeak;
    frontal_info{n, 12} = dChestARPeak;
    frontal_info{n, 13} = pHeadAXPeak;
    frontal_info{n, 14} = pHeadAYPeak;
    frontal_info{n, 15} = pHeadAZPeak;
    frontal_info{n, 16} = pHeadARPeak;
    frontal_info{n, 17} = pChestAXPeak;
    frontal_info{n, 18} = pChestAYPeak;
    frontal_info{n, 19} = pChestAZPeak;
    frontal_info{n, 20} = pChestARPeak;
    frontal_info{n, 21} = pPelvisAXPeak;
    frontal_info{n, 22} = pPelvisAYPeak;
    frontal_info{n, 23} = pPelvisAZPeak;
    frontal_info{n, 24} = pPelvisARPeak;
end
save('IIHS_diadem_dp.mat', 'frontal_info');
close all
