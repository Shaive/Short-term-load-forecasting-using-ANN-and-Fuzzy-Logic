clc
clear all
display('======== Welcome to short term Load forecasting =======');
%% import files
x = input('Do you want to update the database from Excel files in folder (Data) or load existing ones? \n(d == load , u ==Update): ','s');
if x=='u' || x=='U' || ~exist('Data\data.mat', 'file')  || ~exist('Data\valdata.mat', 'file' )
    if x=='d' || x=='D' || ~exist('Data\data.mat', 'file')  || ~exist('Data\valdata.mat', 'file' )
      display('No saved data was found! data will be imported from Excel sheets');
    end
    
    display('data is being imported, please wait.....');
    %%% import the data files
    fileToRead='Data\data.xlsx';      %% traing data 
    fileToRead2='Data\valdata.xlsx';  %% valdation data

    [Load, string] = xlsread(fileToRead, 'Sheet1');
    Load=Load';
    i=1; 
    % convert the string data type dates into numbers year,month,day and hour
    for j = 1: length(string)
        temp=char(string(j));
        year(j) = str2double(temp(2:5));      
        month(j)= str2double(temp(7:8));
        day(j)  = str2double(temp(10:11));
        hour(j) = i;
        i=i+1;
        if(i>24) 
            i=1; 
        end
    end

    dayOfWeek           = weekday(string)';               
    preWeekSameHourLoad = [NaN(1,168), Load(1:end-168)];  
    preDaySameHourLoad  = [NaN(1,24), Load(1:end-24)];    
    pre24HourAverLoad   = filter(ones(1,24)/24, 1, Load); 

    %%% import the valdata file
    [vLoad, string] = xlsread(fileToRead2, 'Sheet1');
    vLoad=vLoad';
    i=1;
    % convert the string data type dates into numbers year,month,day and hour
    for j = 1: length(string)
        temp=char(string(j));
        vyear(j) = str2double(temp(2:5));      
        vmonth(j)= str2double(temp(7:8));
        vday(j)  = str2double(temp(10:11));
        vhour(j) = i;
        i=i+1;
        if(i>24) 
            i=1; 
        end
    end

    vdayOfWeek           = weekday(string)';                               %day Of the Week
    vpreWeekSameHourLoad = [Load(1,(end-191):(end-168))];                 % previous Week Same Hour Load
    vpreDaySameHourLoad  = [Load(1,end-23:end)];                          % previous day same hour Load
    vpre24HourAverLoad    = filter(ones(1,24)/24, 1, Load(1,end-23:end));  % previous 24 Hour Average Load
    display('saving imported data from Excel sheets .....'); 
    save('Data\data','Load','year','month','day','dayOfWeek','hour',....
     'preDaySameHourLoad','preWeekSameHourLoad','pre24HourAverLoad');
 
    save('Data\valdata','vLoad','vyear','vmonth','vday','vdayOfWeek','vhour',...
     'vpreDaySameHourLoad','vpreWeekSameHourLoad','vpre24HourAverLoad');
else 
    load Data\data.mat
    load Data\valdata.mat
end

% remove the temporary variables used
clear x i j temp  string fileToRead fileToRead2;
display('done importing');

%% Initial Exploration of Data
x=input('Do you want to visualize the data?(y == yes else ==no): ','s');
if (x=='y' || x=='Y') 
  
    % power as a function of the time of day.
    figure;
    plot(hour, Load, '.');
    title('Power VS Hour','Fontsize', 12,'color','r'); 
    xlabel('Hour'); ylabel('Power');
    
    % We can see a general trend throughout the day, but we can also clearly
    % see that there is a wide variation throughout the year. Let's use a
    % boxplot to look at it.

    % *Power vs. Hour*
    figure;
    boxplot(Load, hour);
    title('Power VS Hour (boxplot)','Fontsize', 12,'color','r');     
    xlabel('Hour'); ylabel('Power');

    % We can do the same for the other time scales.
    % *Power vs. Day*
    % In general we don't see much variation of power usage based on the day of
    % the month.
    figure;
    boxplot(Load, day, 'plotstyle', 'compact');
    title('Power VS Day of the month','Fontsize', 12,'color','r');         
    xlabel('Day'); ylabel('Power');

    % *Power vs. Day of Week*
    dayText =  {'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'};
    figure;
    boxplot(Load, dayOfWeek,'notch', 'on', 'labels', dayText(unique(dayOfWeek)));
    title('Power VS Day of the week','Fontsize', 12,'color','r');  
    xlabel('Day of Week'); ylabel('Power');

    % *Power vs. Month*
    figure;
    boxplot(Load, month,'notch', 'on', 'labels', month);
    title('Power VS Month','Fontsize', 12,'color','r');      
    xlabel('Month'); ylabel('Power');
end


%% forecasting models menu
again='y';
while (again =='y' || again=='Y')
    run('NN_model');
end