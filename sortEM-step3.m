clear;
data = xlsread('RBL-R T8-3 FR CR CA1 SLM','data');
col = 11;%column number that will use as reference
spine = data(:,col)==1;%pull data from colomn number from all rows, and make it as spine. 

[index, ~] = find(data(:,col)==1);%find data in the data equals 1 and record their location as index. 
data(isnan(data)) = 0;%make the data without 1(text and empty) as 0. 
withmito = (data(index,3));

s1 = sum(withmito);
s2 = sum(spine);

