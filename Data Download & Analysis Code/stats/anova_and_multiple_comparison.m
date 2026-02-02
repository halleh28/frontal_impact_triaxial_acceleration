%for tests by vehicle type

close all
cd 'C:\Users\holzb\Documents\IIHS Research\results';

f = "Stats.xlsx";
data = readtable(f, 'Sheet', 2, 'Range', 'A1:D247'); %change sheet # to perform with different data
c = data.Car;
s = data.SUV;
m = data.Minivan;
p = data.Pickup;

allData = [c s m p];

[a, tbl, stats] = anova1(allData);
figure
multcompare(stats, 'CType', 'bonferroni'); %default is Tukey-Kramer