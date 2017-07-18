% Reads Input Excel File and Generates Machine Part Matrix
inputFileName = 'PFASTInputFile.xlsx';

%% Part Related data
PartDataSheet = 1;
[num,txt,rawPartData] = xlsread(inputFileName,PartDataSheet);
partNum=rawPartData(:,1);
machineSeqOfOperation=rawPartData(:,2);
batchQuantity=rawPartData(:,3);
partNum=partNum(2:end);
machineSeqOfOperation=machineSeqOfOperation(2:end);
batchQuantity=batchQuantity(2:end);
partNumMat = cell2mat(partNum);
batchQuantityMat = cell2mat(batchQuantity);
partNums= unique(partNumMat);
assert(numel(partNums) == numel(partNumMat),'Input Part Data Contains Duplicates');

%% Machine Related data
MachineDatasheet = 2;
numMachineData = xlsread(inputFileName,MachineDatasheet);
machineNumMat = numMachineData(:,1);
machineNums= unique(machineNumMat);
machineNums = sort(machineNums);
assert(numel(machineNumMat) == numel(machineNums),'Input Machine Data Contains Duplicates');

%%  Machine Part Matrix Generation
machinePartMat= zeros(numel(partNums),numel(machineNums));
%operationSequences = cellfun(@(str) regexprep(str,',',' '), machineSeqOfOperation, 'UniformOutput', false);

for indM = 1: numel(partNums)
    tmpMachineStr = machineSeqOfOperation{indM};
    
    if isempty(tmpMachineStr)
         continue
    elseif numel(tmpMachineStr)> 1
        tmpMachineStr= regexprep(tmpMachineStr,',',' ');
        tmpMachinemat= str2num(tmpMachineStr);
        machinePartMat(indM,:) = ismember(machineNums,tmpMachinemat)';
    elseif(isfinite(tmpMachineStr))
        machinePartMat(indM,:) = ismember(machineNums,tmpMachineStr)'; 
    else
        continue
    end
end

%% Hierarchical Clustering
partsDistMat=zeros(numel(partNums));
machineDistMat=zeros(numel(machineNums));
for indP1= 1: numel(partNums)
    tmpPart1= machinePartMat(indP1,:);
    numPart1= sum(tmpPart1);
    for indP2= 1: numel(partNums)
         tmpPart2= machinePartMat(indP2,:);
         numPart2= sum(tmpPart2);
         if (indP1 ==indP2)
             continue
         else
             numCommonMatch= sum(tmpPart2& tmpPart1);
             partsDistMat(indP1,indP2)=(1- (numCommonMatch)/(numPart2+numPart1-numCommonMatch));
             %numCommonMach/numel(machineNums);
         end
    end
end
Y= squareform(partsDistMat);
Z = linkage(Y,'complete');
dendrogram(Z);
%% K-means Clustering
[idx,C,sumd,D] = kmeans(machinePartMat,4,...
    'Display','final','Replicates',5, 'Start','sample');
Clust1= find(idx==1);
Clust2= find(idx==2);
Clust3= find(idx==3);
Clust4= find(idx==4);
%%  Rank Order Clustering (ROC)
%Decimal Num equivalent for each Row
Rows=  1: numel(machineNums);
Rows = sort (Rows, 'descend');
rowsRank = arrayfun(@(row) sum(Rows.*row), machinePartMat);
%Decimal Num equivalent for each Column
Cols=  1: numel(partNums);
Cols = sort (Cols, 'descend');
