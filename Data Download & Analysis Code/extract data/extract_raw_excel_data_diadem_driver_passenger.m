clear
%formatted for data in "diadem\driver + passenger" files

files = dir('C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\driver + passenger\*.xls*');
cd 'C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\driver + passenger';
frontal_info = cell(length(files), 36);

for n = 180:length(files)
    f = files(n).name %'CEF2103.xls';
    [~, sheetNames] = xlsfinfo(f);
%% Test & vehicle data

    testInfo = readtable(f, 'Sheet', 1, 'Range', 'A2:B2', 'ReadVariableNames', false);
    testID = char(string(testInfo.Var1(1)));
    vehInfo = char(string(testInfo.Var2(1)));

    peakData = readtable(f, 'Sheet', 1, 'VariableNamingRule', 'preserve', 'ReadRowNames', true);
    dTestData = readtable(f, 'Sheet', '11_Head_Neck', 'VariableNamingRule', 'preserve');
    pNeckTestData = readtable(f, 'Sheet', '14_Head_Neck', 'VariableNamingRule', 'preserve');
    pLumbarTestData = readtable(f, 'Sheet', '14_Spine', 'VariableNamingRule', 'preserve');
    dRowNamesHead = {'11HEADCG00__ACXA', '11HEADCG00__ACYA', '11HEADCG00__ACZA', '11HEADCG00__ACR'};
    dRowNamesChest = {'11CHST0000__ACXC', '11CHST0000__ACYC', '11CHST0000__ACZC', '11CHST0000__ACRC'};
    dRowNamesNeck = {'11NECKUP00__FOXA', '11NECKUP00__FOZA'};
    pRowNamesHead = {'14HEAD0000HFACXA', '14HEAD0000HFACYA', '14HEAD0000HFACZA', '14HEAD0000HFACR'};
    pRowNamesChest = {'14CHST0000HFACXC', '14CHST0000HFACYC', '14CHST0000HFACZC', '14CHST0000HFACRC'};
    pRowNamesPelvis = {'14PELV0000HFACXA', '14PELV0000HFACYA', '14PELV0000HFACZA', '14PELV0000HFACR'};
    pRowNamesNeck = {'14NECKUP00HFFOXA', '14NECKUP00HFFOZA'};
    pRowNamesLumbar = {'14LUSP0000HFFOXB', '14LUSP0000HFFOZB'};
    vehRowNames = {'10VEHC0000__ACRD', 'VEHC0000__VEX'};
    vehTrunkRowNames = {'10VEHCTRNK__ACRD', 'VEHCTRNK__VEX'};

    dRowNames = [dRowNamesHead, dRowNamesChest, dRowNamesNeck];
    pRowNames = [pRowNamesHead, pRowNamesChest, pRowNamesPelvis, pRowNamesNeck, pRowNamesLumbar];

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
%% Driver neck load cell data

    dNeckFXPeak = max(abs(peakData{dRowNames{9}, 4}), abs(peakData{dRowNames{9}, 5}));
    dNeckFZPeakC = abs(peakData{dRowNames{10}, 4});
    dNeckFZPeakT = abs(peakData{dRowNames{10}, 5});
    dNeckFX = dTestData.(dRowNames{9});
    dNeckFZ = dTestData.(dRowNames{10});

    validRows = ~isnan(dNeckFX) & ~isnan(dNeckFZ);
    
    dNeckFX = dNeckFX(validRows);
    dNeckFZ = dNeckFZ(validRows);
    
    dNeckFRPeak = max(sqrt(dNeckFX.^2 + dNeckFZ.^2));
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
%% Passenger neck load cell data

    pNeckFXPeak = max(abs(peakData{pRowNames{13}, 4}), abs(peakData{pRowNames{13}, 5}));
    pNeckFZPeakC = abs(peakData{pRowNames{14}, 4});
    pNeckFZPeakT = abs(peakData{pRowNames{14}, 5});
    pNeckFX = pNeckTestData.(pRowNames{13});
    pNeckFZ = pNeckTestData.(pRowNames{14});

    validRows = ~isnan(pNeckFX) & ~isnan(pNeckFZ);
    
    pNeckFX = pNeckFX(validRows);
    pNeckFZ = pNeckFZ(validRows);
    
    pNeckFRPeak = max(sqrt(pNeckFX.^2 + pNeckFZ.^2));
%% Passenger lumbar load cell data

    pLumbarFXPeak = max(abs(peakData{pRowNames{15}, 4}), abs(peakData{pRowNames{15}, 5}));
    pLumbarFZPeakC = abs(peakData{pRowNames{16}, 4});
    pLumbarFZPeakT = abs(peakData{pRowNames{16}, 5});
    if isnan(pLumbarFXPeak) || isnan(pLumbarFZPeakT)
        pLumbarFRPeak = '';
    else
        pLumbarFX = pLumbarTestData.(pRowNames{15});
        pLumbarFZ = pLumbarTestData.(pRowNames{16});
        validRows = ~isnan(pLumbarFX) & ~isnan(pLumbarFZ);
        
        pLumbarFX = pLumbarFX(validRows);
        pLumbarFZ = pLumbarFZ(validRows);
        
        pLumbarFRPeak = max(sqrt(pLumbarFX.^2 + pLumbarFZ.^2));
    end
%% add all data to cell array (for export to excel later)

    r = n;
    r = r - 179;
    frontal_info{r, 1} = testID;
    frontal_info{r, 2} = vehInfo;
    frontal_info{r, 3} = vehdvPeak;
    frontal_info{r, 4} = vehAPeak;
    frontal_info{r, 5} = dHeadAXPeak;
    frontal_info{r, 6} = dHeadAYPeak;
    frontal_info{r, 7} = dHeadAZPeak;
    frontal_info{r, 8} = dHeadARPeak;
    frontal_info{r, 9} = dChestAXPeak;
    frontal_info{r, 10} = dChestAYPeak;
    frontal_info{r, 11} = dChestAZPeak;
    frontal_info{r, 12} = dChestARPeak;
    frontal_info{r, 13} = dNeckFXPeak;
    frontal_info{r, 14} = dNeckFZPeakC;
    frontal_info{r, 15} = dNeckFZPeakT;
    frontal_info{r, 16} = dNeckFRPeak;
    frontal_info{r, 17} = pHeadAXPeak;
    frontal_info{r, 18} = pHeadAYPeak;
    frontal_info{r, 19} = pHeadAZPeak;
    frontal_info{r, 20} = pHeadARPeak;
    frontal_info{r, 21} = pChestAXPeak;
    frontal_info{r, 22} = pChestAYPeak;
    frontal_info{r, 23} = pChestAZPeak;
    frontal_info{r, 24} = pChestARPeak;
    frontal_info{r, 25} = pPelvisAXPeak;
    frontal_info{r, 26} = pPelvisAYPeak;
    frontal_info{r, 27} = pPelvisAZPeak;
    frontal_info{r, 28} = pPelvisARPeak;
    frontal_info{r, 29} = pNeckFXPeak;
    frontal_info{r, 30} = pNeckFZPeakC;
    frontal_info{r, 31} = pNeckFZPeakT;
    frontal_info{r, 32} = pNeckFRPeak;
    frontal_info{r, 33} = pLumbarFXPeak;
    frontal_info{r, 34} = pLumbarFZPeakC;
    frontal_info{r, 35} = pLumbarFZPeakT;
    frontal_info{r, 36} = pLumbarFRPeak;
end
save('IIHS_diadem_dp.mat', 'frontal_info');
close all