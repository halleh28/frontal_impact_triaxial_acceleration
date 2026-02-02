clear
%formatted for data in "plots" files

files = dir('C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\plots\*.xls*');
cd 'C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\plots';
frontal_info = {length(files), 12};

for n = 1:length(files)
    f = files(n).name; %'CEF0101plots.xls';

    %% Test & vehicle data
    testInfo = readtable(f, 'Sheet', 2, 'Range', 'A1:A2', 'ReadVariableNames', false);
    testID = char(string(testInfo.Var1(1)));
    vehInfo = char(string(testInfo.Var1(2)));
    
    testData = readtable(f, 'Sheet', 2, 'Range', 'A6:BB6000', 'ReadVariableNames', false); %extra rows included for potentially longer data sets

    %peak vehicle a and dV (dV calculated based on spreadsheet)
    vehA = testData.Var5;
    vehA = vehA(~isnan(vehA));
    vehAPeak = max(abs(vehA));
    vehdv = testData.Var6;
    vehdv = vehdv(~isnan(vehdv));
    vehdvPeak = max(vehdv) - min(vehdv);

    %% Head accel data
    headAX = testData.Var7;
    headAX = headAX(~isnan(headAX));
    headAXPeak = max(abs(headAX));

    headAY = testData.Var8;
    headAY = headAY(~isnan(headAY));
    headAYPeak = max(abs(headAY));

    headAZ = testData.Var9;
    headAZ = headAZ(~isnan(headAZ));
    headAZPeak = max(abs(headAZ));

    headAR = testData.Var10;
    headAR = headAR(~isnan(headAR));
    headARPeak = max(abs(headAR));

    %% Chest accel data
    chestAX = testData.Var19;
    chestAX = chestAX(~isnan(chestAX));
    chestAXPeak = max(abs(chestAX));

    chestAY = testData.Var20;
    chestAY = chestAY(~isnan(chestAY));
    chestAYPeak = max(abs(chestAY));

    chestAZ = testData.Var21;
    chestAZ = chestAZ(~isnan(chestAZ));
    chestAZPeak = max(abs(chestAZ));

    chestAR = testData.Var22;
    chestAR = chestAR(~isnan(chestAR));
    chestARPeak = max(abs(chestAR));

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
save('IIHS_plots.mat', 'frontal_info');
close all
