clear
%formatted for data in "plots" files

files = dir('C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\plots\*.xls*');
cd 'C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\plots';
frontal_info = {length(files), 12};

for n = 1:length(files)
    n
    f = files(n).name; %'CEF0101plots.xls';

    testInfo = readtable(f, 'Sheet', 2, 'Range', 'A1:A2', 'ReadVariableNames', false);
    testID = char(string(testInfo.Var1(1)));

    accelType = readtable(f, 'Sheet', 2, 'Range', 'A4:Z4', 'ReadVariableNames', false, 'VariableNamingRule', 'preserve', 'TextType', 'string');
    
    testData = readtable(f, 'Sheet', 2, 'Range', 'A5:BB6000', 'VariableNamingRule', 'preserve');
    
    out_file = 'combining test data macro plots.xlsm';
%% Vehicle accel data

    veh_idx = find(strcmpi(accelType{1, :}, 'VEHICLE'), 1);
    var_names = testData.Properties.VariableNames;
    veh_cols = [1, veh_idx:veh_idx+3];
    veh_block = make_block(testID, testData, veh_cols);

    block_width = size(veh_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(veh_block, out_file, 'Sheet', 'Vehicle', 'Range', range_str);
%% Head accel data

    head_idx = find(strcmpi(accelType{1, :}, 'HEAD'), 1);
    head_cols = [1, head_idx:head_idx+3];
    head_block = make_block(testID, testData, head_cols);

    block_width = size(head_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(head_block, out_file, 'Sheet', 'Head', 'Range', range_str);
%% Chest accel data

    chest_idx = find(strcmpi(accelType{1, :}, 'CHEST'), 1);
    chest_cols = [1, chest_idx+1:chest_idx+4];
    chest_block = make_block(testID, testData, chest_cols);

    block_width = size(chest_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(chest_block, out_file, 'Sheet', 'Chest', 'Range', range_str);
end
%%
function block = make_block(testID, testData, col_idx)
    n_rows = height(testData) + 2;
    var_names = testData.Properties.VariableNames;

    block = cell(n_rows, numel(col_idx));

    block{1,1} = testID;

    block(2, :) = var_names(col_idx);

    block(3:end, :) = table2cell(testData(:, col_idx));
end

function col_str = num2xlcol(col_num)
% Preallocate 3 characters (max Excel column length)
    buf = repmat('A', 1, 3);  
    k = 3;

    while col_num > 0
        col_num = col_num - 1;
        buf(k) = char('A' + mod(col_num, 26));
        col_num = floor(col_num / 26);
        k = k - 1;
    end

    col_str = string(buf(k+1:end));
end