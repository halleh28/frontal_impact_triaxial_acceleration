clear
%formatted for data in "diadem\driver + passenger" files

files = dir('C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\driver + passenger\*.xls*');
cd 'C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\driver + passenger';

for n = 1:length(files)    
    n
    f = files(n).name; %'CEF2103.xls';
    [~, sheet_names] = xlsfinfo(f);

    testInfo = readtable(f, 'Sheet', 1, 'Range', 'A2:R2', 'ReadVariableNames', false);
    testID = char(string(testInfo.Var1(1)));
    
    vehicle = readtable(f, 'Sheet', find(contains(sheet_names, 'Vehicle')), 'VariableNamingRule', 'preserve');
    d_head = readtable(f, 'Sheet', '11_Head_Neck', 'VariableNamingRule', 'preserve');
    d_chest = readtable(f, 'Sheet', '11_Chest', 'VariableNamingRule', 'preserve');
    p_head = readtable(f, 'Sheet', '14_Head_Neck', 'VariableNamingRule', 'preserve');
    p_chest = readtable(f, 'Sheet', '14_Chest', 'VariableNamingRule', 'preserve');
    p_pelvis = readtable(f, 'Sheet', '14_Femur_Pelvis', 'VariableNamingRule', 'preserve');

    d_head_row_names = {'11HEADCG00__ACXA', '11HEADCG00__ACYA', '11HEADCG00__ACZA', '11HEADCG00__ACR'};
    d_chest_row_names = {'11CHST0000__ACXC', '11CHST0000__ACYC', '11CHST0000__ACZC', '11CHST0000__ACRC'};
    p_head_row_names = {'14HEAD0000HFACXA', '14HEAD0000HFACYA', '14HEAD0000HFACZA', '14HEAD0000HFACR'};
    p_chest_row_names = {'14CHST0000HFACXC', '14CHST0000HFACYC', '14CHST0000HFACZC', '14CHST0000HFACRC'};
    p_pelvis_row_names = {'14PELV0000HFACXA', '14PELV0000HFACYA', '14PELV0000HFACZA', '14PELV0000HFACR'};
    veh_row_names = {'10VEHC0000__ACXD', '10VEHC0000__ACYD', '10VEHC0000__ACZD', '10VEHC0000__ACRD'};
    veh_trunk_row_names = {'10VEHCTRNK__ACXD_NoName*', '10VEHCTRNK__ACXD', '10VEHCTRNK__ACYD', '10VEHCTRNK__ACZD', '10VEHCTRNK__ACRD'};

    out_file = 'combining test data macro diadem dp.xlsm';
%% Vehicle accel data

    veh_var_names = vehicle.Properties.VariableNames;

    if isnumeric(vehicle{1, veh_row_names{1}}) && isscalar(vehicle{1, veh_row_names{1}}) && ~isnan(vehicle{1, veh_row_names{1}})
        veh_idx = find(ismember(veh_var_names, veh_row_names));
        veh_cols  = [1, veh_idx];
    else
        veh_idx = find(ismember(veh_var_names, veh_trunk_row_names));
        veh_cols = [veh_idx];
    end

    veh_block = make_block(testID, vehicle, veh_cols, veh_var_names);

    block_width = size(veh_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);
    
    writecell(veh_block, out_file, 'Sheet', 'Vehicle', 'Range', range_str);
%% Driver head accel data

    d_head_var_names = d_head.Properties.VariableNames;
    d_head_idx = find(ismember(d_head_var_names, d_head_row_names));
    d_head_cols = [1, d_head_idx];
    d_head_block = make_block(testID, d_head, d_head_cols, d_head_var_names);

    block_width = size(d_head_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(d_head_block, out_file, 'Sheet', 'Driver Head', 'Range', range_str);
%% Driver chest accel data

    d_chest_var_names = d_chest.Properties.VariableNames;
    d_chest_idx = find(ismember(d_chest_var_names, d_chest_row_names));
    d_chest_cols = [1, d_chest_idx];
    d_chest_block = make_block(testID, d_chest, d_chest_cols, d_chest_var_names);

    block_width = size(d_chest_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(d_chest_block, out_file, 'Sheet', 'Driver Chest', 'Range', range_str);
%% Passenger head accel data

    p_head_var_names = p_head.Properties.VariableNames;
    p_head_idx = find(ismember(p_head_var_names, p_head_row_names));
    p_head_cols = [1, p_head_idx];
    p_head_block = make_block(testID, p_head, p_head_cols, p_head_var_names);

    block_width = size(p_head_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(p_head_block, out_file, 'Sheet', 'Passenger Head', 'Range', range_str);
%% Passenger chest accel data

    p_chest_var_names = p_chest.Properties.VariableNames;
    p_chest_idx = find(ismember(p_chest_var_names, p_chest_row_names));
    p_chest_cols = [1, p_chest_idx];
    p_chest_block = make_block(testID, p_chest, p_chest_cols, p_chest_var_names);

    block_width = size(p_chest_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(p_chest_block, out_file, 'Sheet', 'Passenger Chest', 'Range', range_str);
%% Passenger pelvis accel data

    p_pelvis_var_names = p_pelvis.Properties.VariableNames;
    p_pelvis_idx = find(ismember(p_pelvis_var_names, p_pelvis_row_names));
    p_pelvis_cols = [1, p_pelvis_idx];
    p_pelvis_block = make_block(testID, p_pelvis, p_pelvis_cols, p_pelvis_var_names);

    block_width = size(p_pelvis_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(p_pelvis_block, out_file, 'Sheet', 'Passenger Pelvis', 'Range', range_str);
end
%%
function block = make_block(testID, testData, col_idx, var_names)
    n_rows = height(testData) + 2;

    block = cell(n_rows, numel(col_idx));

    block{1,1} = testID;

    block(2, :) = var_names(col_idx);

    block(3:end, :) = table2cell(testData(:, col_idx));
end

function col_str = num2xlcol(col_num)
%NUM2XLCOL Convert 1-based column index to Excel-style column letters.
    buf = repmat('A', 1, 3);
    k = 3;

    n = col_num;
    while n > 0
        n = n - 1;
        buf(k) = char('A' + mod(n, 26));
        n = floor(n / 26);
        k = k - 1;
    end

    col_str = string(buf(k+1:end));
end