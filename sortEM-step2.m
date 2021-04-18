%fileToRead1 = 'RBL-R T8-3 FR CR CA3 SLM LX112 B1-200sec-1 ROI00 traceLists';
outFile = 'TestTracelist';
rowN = 0;

allfiles = dir('Results/*traceLists');
Nfiles = length({allfiles.name});
outputdata = [];
flist = {};
for fileN = 1:Nfiles
    
    fileToRead1 = ['Results\',allfiles(fileN).name]
    
    newData1 = importdata(fileToRead1);
    
    
    data = newData1.data;
    
    textdata = newData1.textdata;
    
    
    %cellfun(@(IDX) ~isempty(IDX), strfind(textdata(:,1), 'CR'))
    
    Ndata = length(data);
    
    %% Get all bouton numbers
    boutonNs =[];
    for i = 1:Ndata
        temp = char(textdata(i,2));
        allNs = regexp(temp,'\d*','Match');
        
        if ~isempty(allNs)
            boutonNs = [boutonNs allNs];
        end
        
    end
    
    boutonNs = unique(boutonNs)
    
    Nboutons = length(boutonNs);
    
    %%
    
    for i = 1:Nboutons
        rowN = rowN + 1;
        idxN = cellfun(@(IDX) ~isempty(IDX), strfind(textdata(:,2), boutonNs(i)));
        idxB = cellfun(@(IDX) ~isempty(IDX), strfind(textdata(:,2), '-B'));
        idxsyn = cellfun(@(IDX) ~isempty(IDX), strfind(textdata(:,2), '-syn'));
        
        boutonB = idxN&idxB;
        boutonsyn = idxN&idxsyn;
        
        sectionB = textdata(boutonB,1);
        sectionsyn = textdata(boutonsyn,1);
        
        [C,iB,isyn] = intersect(sectionB,sectionsyn);
        
        d1 = data(boutonB(2:end),end);
        d2 = d1(iB);
        maxd = max(d2);
        
        
        outputdata = [outputdata; [str2num(char(boutonNs(i))) maxd]];
        %idx = zeros(length(boutonsyn));
        flist{end+1} = fileToRead1;
        %     for j = 1:length(sectionno)
        %
        %     end
        
    end
end

Tab = table(outputdata,flist');

xlswrite(outFile,{'Bouton #'},1,'A1');
xlswrite(outFile,{'Diameter'},1,'B1');
xlswrite(outFile,{'Filename'},1,'C1');
xlswrite(outFile,outputdata,1,'A2');
xlswrite(outFile,flist',1,'C2');