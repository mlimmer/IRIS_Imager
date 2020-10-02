% Read in Munsel colors from scanned in color swatch and ouput CIELAB values
% The user will select the center of the square and the code will average a
% block of pixels
clear
clc
close all
%% Constants
DPI=200; %Scanner resolution
Box_halfwidth=0.15*DPI; %inches of color swatch->pixels

%% Prompts
ImFile=input('Enter file name for Munsell Color Sheet including extension\n','s');

%% Image Import and Clean
I_Munsell=imread(ImFile);
I_Munsell_lab=rgb2lab(I_Munsell,'ColorSpace','srgb');
figure;
imshow(I_Munsell);

%% Vectors
ncell=inputdlg('How many samples are there on this image?');
n=str2double(ncell{:});
Box_Centers=zeros(n,2);
Names=cell(n,1);
Box_RGB=zeros(n,3);
Box_LAB=Box_RGB;


%% Pick Boxes
% figure;
% imshow(I_Munsell)
title('Select center of box and press SPACE');
for i=1:n    
    dcm_obj = datacursormode; %Enable point selection
    set(dcm_obj,'DisplayStyle','datatip',...
        'SnapToDataVertex','off','Enable','on')
    pause; %wait for user to hit a key
    Point = getCursorInfo(dcm_obj); %ID point coordinates
    Box_Centers(i,:) = Point.Position; %ID point coordinates
    % Determine average RGB in boxes
    Box_RGB(i,:)=[mean(mean(I_Munsell(Box_Centers(i,2)-Box_halfwidth:Box_Centers(i,2)+Box_halfwidth,Box_Centers(i,1)-Box_halfwidth:Box_Centers(i,1)+Box_halfwidth,1))),mean(mean(I_Munsell(Box_Centers(i,2)-Box_halfwidth:Box_Centers(i,2)+Box_halfwidth,Box_Centers(i,1)-Box_halfwidth:Box_Centers(i,1)+Box_halfwidth,2))),mean(mean(I_Munsell(Box_Centers(i,2)-Box_halfwidth:Box_Centers(i,2)+Box_halfwidth,Box_Centers(i,1)-Box_halfwidth:Box_Centers(i,1)+Box_halfwidth,3)))];
    Box_LAB(i,:)=[mean(mean(I_Munsell_lab(Box_Centers(i,2)-Box_halfwidth:Box_Centers(i,2)+Box_halfwidth,Box_Centers(i,1)-Box_halfwidth:Box_Centers(i,1)+Box_halfwidth,1))),mean(mean(I_Munsell_lab(Box_Centers(i,2)-Box_halfwidth:Box_Centers(i,2)+Box_halfwidth,Box_Centers(i,1)-Box_halfwidth:Box_Centers(i,1)+Box_halfwidth,2))),mean(mean(I_Munsell_lab(Box_Centers(i,2)-Box_halfwidth:Box_Centers(i,2)+Box_halfwidth,Box_Centers(i,1)-Box_halfwidth:Box_Centers(i,1)+Box_halfwidth,3)))];
    rectangle('Position',[Box_Centers(i,1)-Box_halfwidth, Box_Centers(i,2)-Box_halfwidth, 2*Box_halfwidth, 2*Box_halfwidth],'FaceColor',Box_RGB(i,:)/255);
    Names{i}=inputdlg('Input Color Name (e.g., 10YR5/4)');
    title('Select center of next box and press SPACE');
    
end

%% Output Data
OutFile='Musell_Colors.xlsx';
for i=1:n
    OutData(i,:)=[Names{i},num2cell(Box_LAB(i,:))];
end
xlswrite(OutFile,OutData,ImFile);

