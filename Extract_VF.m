folder = uigetdir();

d = dir(folder);
isub = [d(:).isdir]; %# returns logical vector
nameFolds = {d(isub).name}';
nameFolds(ismember(nameFolds,{'.','..'})) = [];
t = size(nameFolds,1);

b = [];
if t==0 
    tiffiles = dir([folder  '/*.xlsx']); 
      sd = length(tiffiles);
    
    %read table
    for k = 1:sd
     File1 = [folder  '/' tiffiles(k).name];
        A = xlsread(File1);
        table{1,k} = A;
    end

%extract vf readings
    for k = 1:sd
        if isempty(table{1,k})
        else
        md(k) = table{1,k}(66,1);
        psd(k) = table{1,k}(68,1);
        vfi(k) = table{1,k}(71,1);
        end
    end

    md = md(:);
    psd = psd(:);
    vfi = vfi(:);

    vf = [md psd vfi];
    vf = num2cell(vf);

%get pid,vis_date,eye by splitting the Excel file name
    for k = 1:sd
        if isempty(table{1,k})
        else
        c = strsplit(tiffiles(k).name,{'_','.'});
        pid(k) = c(1);
        vis_date(k) = c(2);
        eye(k) =c(3);
        end
    end

    pid =pid(:);
    vis_date = vis_date(:);
    eye = eye(:);

    f = [pid vis_date eye vf];
    

    
    b = [b;f];
else
for i = 1 : t
    f= [];
    table = {};
    md = [];
    psd = [];
    vfi = [];
    pid = {};
    vis_date = {};
    eye = {};
    vf = [];
    s = char(strcat(folder,{'/'},char(nameFolds{i})));
    sd = length(tiffiles);
    
   
    tiffiles = dir([s  '/*.xlsx']);
    sd = length(tiffiles);
    
    %read table
    for k = 1:sd
     File1 = [s  '/' tiffiles(k).name];
        A = xlsread(File1);
        table{1,k} = A;
    end

%extract vf readings
    for k = 1:sd
        if isempty(table{1,k})
        else
        md(k) = table{1,k}(66,1);
        psd(k) = table{1,k}(68,1);
        vfi(k) = table{1,k}(71,1);
        end
    end

    md = md(:);
    psd = psd(:);
    vfi = vfi(:);

    vf = [md psd vfi];
    vf = num2cell(vf);

%get pid,vis_date,eye by splitting the Excel file name
    for k = 1:sd
        if isempty(table{1,k})
        else
        c = strsplit(tiffiles(k).name,{'_','.'});
        pid(k) = c(1);
        vis_date(k) = c(2);
        eye(k) =c(3);
        end
    end

    pid =pid(:);
    vis_date = vis_date(:);
    eye = eye(:);

    f = [pid vis_date eye vf];
    

    
    b = [b;f];
end
end
 
filename = 'VF.xlsx';
xlswrite(filename,b);
