clear
%formatted for data in "diadem\passenger" files

files = dir('C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\passenger\*.xls*');
cd 'C:\Users\holzb\Documents\IIHS Research\Data (IIHS)\diadem\passenger';

for n = 1:length(files)    
    n
    f = files(n).name; %'CF16005 (Passenger).xlsx';
    [~, sheet_names] = xlsfinfo(f);

    testInfo = readtable(f, 'Sheet', 1, 'Range', 'A2:B2', 'ReadVariableNames', false);
    testID = char(string(testInfo.Var1(1)));
    
    head = readtable(f, 'Sheet', 'Head_Neck', 'VariableNamingRule', 'preserve');
    chest = readtable(f, 'Sheet', 'Chest', 'VariableNamingRule', 'preserve');

    head_row_names = {'13HEADCG00__ACXA', '13HEADCG00__ACYA', '13HEADCG00__ACZA', '13HEADCG00__ACR'};
    chest_row_names = {'13CHST0000__ACXC', '13CHST0000__ACYC', '13CHST0000__ACZC', '13CHST0000__ACRC'};

    out_file = 'combining test data macro diadem p.xlsm';
%% Passenger head accel data

    head_var_names = head.Properties.VariableNames;
    head_idx = find(ismember(head_var_names, head_row_names));
    head_cols = [1, head_idx];
    head_block = make_block(testID, head, head_cols, head_var_names);

    block_width = size(head_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(head_block, out_file, 'Sheet', 'Passenger Head', 'Range', range_str);
%% Passenger chest accel data

    chest_var_names = chest.Properties.VariableNames;
    chest_idx = find(ismember(chest_var_names, chest_row_names));
    chest_cols = [1, chest_idx];
    chest_block = make_block(testID, chest, chest_cols, chest_var_names);

    block_width = size(chest_block, 2);
    start_idx = (n-1)*(block_width + 1) + 1;
    start_col = num2xlcol(start_idx);
    range_str = sprintf('%s1', start_col);

    writecell(chest_block, out_file, 'Sheet', 'Passenger Chest', 'Range', range_str);
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
    buf = repmat('A', 1, 1);
    k = 1;

    n = col_num;
    while n > 0
        n = n - 1;
        buf(k) = char('A' + mod(n, 26));
        n = floor(n / 26);
        k = k - 1;
    end

    col_str = string(buf(k+1:end));
end