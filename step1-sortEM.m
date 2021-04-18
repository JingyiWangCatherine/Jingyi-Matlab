outFile = 'allResultsLatest';

sheet  = 1;

allfiles = dir('Results/*objects');
Nfiles = length({allfiles.name});

rowN = 2; % To skip first two rows for placing headers;

for fileN = 1:Nfiles
    fileToRead1 = ['Results\',allfiles(fileN).name];
    disp(fileToRead1);
    newData1 = importdata(fileToRead1);
    data = newData1.data;
    textdata = newData1.textdata;
    sz = size(data,2);
    
    for i=2:length(textdata)
        
        temp1 = textdata(i-1,1);
        temp2 = textdata(i,1);
        
        str1 = char(temp1);
        str2 = char(temp2);
        
        allN1s = regexp(str1,'\d*','Match');
        allNs = regexp(str2,'\d*','Match'); % Extract all numbers in the text string
        boutonN = allNs(1); %Extract the bouton number
%         boutonN
%         str1
%         str2
        if isempty(allN1s)
            allN1s = '0'; %Start afresh for a new file
        end
        
        if str2num(char(allNs(1))) ~= str2num(char(allN1s(1))) & ~isempty(str2num(str2(1)))
            rowN = rowN+1; %move to new row if bouton number changes
            xlswrite(outFile,boutonN,sheet,['A' num2str(rowN)]) ;
            disp(boutonN);
            origrow = rowN;
        end
        
        cc1 = ' on sp CR n';
        cc2 = ' on sp CR p';
        cc3 = ' on Den CR p light';
        cc4 = ' sp CR p light';
        cc5 = ' on Den CR n';
        cc6 = ' on Den CR p';
        cc7 = '-B without mito';
        cc8 = '-B with mito';
        cc9 = '-syn round';
        cc10 = '-syn perforanted';
        cc11 = ' on sp SA n';
        cc12 = ' on sp SA p';
        
        ccc1 = strcat(boutonN,cc1);
        ccc2 = strcat(boutonN,cc2);
        ccc3 = strcat(boutonN,cc3);
        ccc4 = strcat(boutonN,cc4);
        ccc5 = strcat(boutonN,cc5);
        ccc6 = strcat(boutonN,cc6);
        ccc7 = strcat(boutonN,cc7);
        ccc8 = strcat(boutonN,cc8);
        ccc9 = strcat(boutonN,cc9);
        ccc10 = strcat(boutonN,cc10);
        ccc11 = strcat(boutonN,cc11);
        ccc12 = strcat(boutonN,cc12);
        
        if ~isempty(strfind(str2,cc1))
            xlswrite(outFile,1,sheet,['K' num2str(origrow)]);
        elseif ~isempty(strfind(str2,cc2))
            xlswrite(outFile,1,sheet,['G' num2str(origrow)]);
        elseif  ~isempty(strfind(str2,cc3))
            xlswrite(outFile,1,sheet,['H' num2str(origrow)]);
        elseif ~isempty(strfind(str2,cc4))
            xlswrite(outFile,1,sheet,['I' num2str(origrow)]);
        elseif ~isempty(strfind(str2,cc5))
            xlswrite(outFile,1,sheet,['J' num2str(origrow)]);
        elseif ~isempty(strfind(str2,cc6))
            xlswrite(outFile,1,sheet,['F' num2str(origrow)]);
        elseif ~isempty(strfind(str2,cc7))            
                 xlswrite(outFile,1,sheet,['C' num2str(origrow)]);
                 xlswrite(outFile,data(i-1,sz-1),sheet,['M' num2str(origrow)]);
                 xlswrite(outFile,data(i-1,sz),sheet,['N' num2str(origrow)]);
                 %disp(['wrote C ',str2])
        elseif ~isempty(strfind(str2,cc8))            
                 xlswrite(outFile,1,sheet,['B' num2str(origrow)]);
                 xlswrite(outFile,data(i-1,sz-1),sheet,['M' num2str(origrow)]);
                 xlswrite(outFile,data(i-1,sz),sheet,['N' num2str(origrow)]);
                 %disp(['wrote B ',str2])
        elseif ~isempty(strfind(str2,cc9))                       
                 xlswrite(outFile,1,sheet,['E' num2str(origrow)]);
                 xlswrite(outFile,data(i-1,sz-1),sheet,['P' num2str(origrow)]);
         elseif ~isempty(strfind(str2,cc10))                       
                 xlswrite(outFile,1,sheet,['D' num2str(origrow)]);
                 xlswrite(outFile,data(i-1,sz-1),sheet,['P' num2str(origrow)]);
        elseif ~isempty(strfind (str2,cc11))
            xlswrite(outFile,1,sheet,['R' num2str(origrow)]);
        elseif ~isempty(strfind (str2,cc12))
            xlswrite(outFile,1,sheet,['Q' num2str(origrow)]);
        else
            disp(['row #',num2str(i),' is ',char(temp2)]);
        end
        
        
        bN = char(boutonN);
        %[~, idxA, idxB] = intersect(str2,bN);
        idxA = find(ismember(str2,bN));
        
        remtext = str2(idxA(end)+1:end);
        if ~isempty(remtext)
            if (remtext(1) == 'a' || remtext(1) == 'b' || remtext(1) == 'c' || remtext(1) == 'd')
                rowN = rowN + 1;
                if ~isempty(strfind(str2,'syn'))
                    xlswrite(outFile,data(i-1,sz-1),sheet,['P' num2str(rowN)]) ;
                end
                
                if ~isempty(strfind(str2,'round'))
                    
                    %xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['E' num2str(rowN)]) ;
                end
                
                if ~isempty(strfind(str2,'perforanted'))
                    
                    %xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['D' num2str(rowN)]) ;
                end
                
                
                if ~isempty(strfind(str2,cc6))
                    
                    xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['F' num2str(rowN)]) ;
                end
                
                if ~isempty(strfind(str2,cc2)) && isempty(strfind(str2,cc4))
                    %rowN = rowN + 1;
                    xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['G' num2str(rowN)]) ;
                end
                
                if ~isempty(strfind(str2,cc3))
                    %rowN = rowN + 1;
                    xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['H' num2str(rowN)]) ;
                end
                
                if ~isempty(strfind(str2,cc4))
                    %rowN = rowN + 1;
                    xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['I' num2str(rowN)]) ;
                end
                
                if ~isempty(strfind(str2,cc5))
                    %rowN = rowN + 1;
                    xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['J' num2str(rowN)]) ;
                end
                
                if ~isempty(strfind(str2,cc1))
                    %rowN = rowN + 1;
                    xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['K' num2str(rowN)]) ;
                end
                
                if ~isempty(strfind(str2,'multi'))
                    %rowN = rowN + 1;
                    %xlswrite(outFile,{str2},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,{str2(1:idxA(end)+1)},sheet,['A' num2str(rowN)]) ;
                    xlswrite(outFile,1,sheet,['L' num2str(rowN)]) ;
                end
                
                
            end
        end
        
        
        %
    end
end
disp(['File generated: ',outFile]);