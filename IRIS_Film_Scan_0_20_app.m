%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% IRIS Imager - IRIS Film Image Analysis
% MAL 10/10/18
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

clc
close all


%% Version Control
% 0.1 - Initial Version
% 0.2 - Change from looking for cirlces to looking for corners of rectangle
% 0.3 - All corner hunting is now done manually
% 0.4 - Corrections after using real films, additional HSV image processing
% 0.5 - Switched to CIEL*A*B* colorspace to avoid confounding S & V in HSV
% 0.6 - Added circle masking
% 0.7 - Added clustering and Munsell Color Matching
% 0.8 - Removed one circle and rotation for 2019 films Abby Evans 10/30/19
% 0.9 - Changed circle hunting for black holes instead of white holes,
% added an output for censoring removal % by 0
% 0.10 - Added output to test using RGB for removal
% 0.11 - Screens out Fe from Mn films and loads existing saved files
% 0.12 - Pull in constants from excel
% 0.13 - Added to Mn film output
% 0.14 - Added option to output removal with depth
% 0.15 - Moved to switch case structure for different film types
% 0.16 - Added sulfide code
% 0.17 - Changed file input mode
% 0.18 - Changed pausing for corner selection
% 0.18app - Changed output to output to app screen
% 0.19app - Changed xls functions to table functions
% 0.20app - Fixed bug in call to CIRCLE_HUNTER where input radii need to be
% rounded

try
%% Constants - read from file
Messages={};
Messages{1,1}=sprintf('Loading parameters...\n');
app.OutputText.Value=Messages;
ParamFile='AdjustableParameters.xlsx'; %Parameter file name
Parameters=readtable(app.Inputfile.Value,'Sheet','Parameters','Range','B1:B25');

DPI=str2double(Parameters.Value{1}); %image resolution in pixels/inch
x_Crop=str2double(Parameters.Value{2}); %number of pixels to crop off x direction
y_Crop=str2double(Parameters.Value{3}); %number of pixels to crop off y direction
Scale_Factor=str2double(Parameters.Value{4}); %Ratio to scale the image by averaging pixels i.e., 0.5 = 1/2 size
Start_Upside_down_flag=str2double(Parameters.Value{5}); %Flag for upsidedown scanned in images 1=upsidedown
End_Upside_down_flag=str2double(Parameters.Value{6}); %Flag for upsidedown scanned in images 1=upsidedown
White_L=str2double(Parameters.Value{7}); %Lightness value for white film (LAB)
Fe_b_Threshold=str2double(Parameters.Value{8}); %Values of b* change above this are considered Fe (rather than white) for counting Fe on Mn films
CircleRad_min=str2double(Parameters.Value{9})*DPI*Scale_Factor; %Min cirlce radius in pixels;
CircleRad_max=str2double(Parameters.Value{10})*DPI*Scale_Factor; %Max cirlce radius in pixels;
Crop_Buffer=str2double(Parameters.Value{11}); %Factor for cropping circles to increase image area by a factor of 2
Circle_NaN_buffer=str2double(Parameters.Value{12}); %Factor for increasing NaN masking around circles (multiplied by circle radius)
n_circles_start=str2double(Parameters.Value{13}); %number of circles in starting image
n_circles_end=str2double(Parameters.Value{14}); %number of circles in ending image
n_clusters_start=str2double(Parameters.Value{15}); %number of clusters to find in starting image
n_clusters_end=str2double(Parameters.Value{16}); %number of clusters to find in ending image
OutputRemoval=str2double(Parameters.Value{17}); %flag: 1 to output removal with depth, 0 to supress output
OutputRemovalDepth=str2double(Parameters.Value{18}); %Distance over which to average removal with depth (inches). Use 0 without averaging
IRIS_Type=Parameters.Value{19}; %Type of IRIS film analyzed -- Fe, Mn or S
OutFile=Parameters.Value{20}; %Output data file name
OutSheet=Parameters.Value{21}; %Output datasheet
Start_circle_color=Parameters.Value{22}; %Color of circle relative to sheet for starting image
End_circle_color=Parameters.Value{23}; %Color of circle relative to sheet for ending image
Image_Path=Parameters.Value{24}; %Default file path for images

%% Prompts
if ~strcmp(IRIS_Type,'S') %Start file not needed for S films
    [Start_file,Start_path] = uigetfile({'*.jpg;*.bmp;*.gif;*.tif;*.png','Image files (*.jpg,*.bmp,*.gif,*.tif,*.png)';...
        '*.*','All files'},'Select initial film',Image_Path);
    if ~ischar(Start_file)
        Messages{size(Messages,1)+1,1}=(sprintf('No start file selected\n'));
        app.OutputText.Value=Messages;
    return;
    end
end

[End_file,End_path] = uigetfile({'*.jpg;*.bmp;*.gif;*.tif;*.png','Image files (*.jpg,*.bmp,*.gif,*.tif,*.png)';...
        '*.*','All files'},'Select final film',Image_Path);

if ~ischar(End_file)
    Messages{size(Messages,1)+1,1}=(sprintf('No end file selected\n'));
        app.OutputText.Value=Messages;
    return;
end

Messages{size(Messages,1)+1,1}=(sprintf('Files selected, checking to see if processed files exist.\n'));
app.OutputText.Value=Messages;
    
    
%% Check to see if output file already exists
filename=[erase(End_file,'.'),'.mat'];
filename=fullfile(End_path,filename);
if exist(filename,'file')==2
    Messages{size(Messages,1)+1,1}=(sprintf('Output file already exists!  Loading output file.\n'));
    app.OutputText.Value=Messages;
     
    load(filename)
else %Load images and align/mask

    End_fullfile=fullfile(End_path,End_file);
    switch IRIS_Type
        case 'S'
            %% Image Import and Clean
            I_end=imread(End_fullfile);
            if End_Upside_down_flag==1
                I_end=imrotate(I_end, 180);
            end

            %% Identify Corners in image
            fighan=figure;
            set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.

            imshow(I_end);
            User_End_Corners=CORNER_ID(fighan);
            End_UL=User_End_Corners(1,:);
            End_UR=User_End_Corners(2,:);
            End_LR=User_End_Corners(3,:);
            End_LL=User_End_Corners(4,:);

            close(fighan);

             %% Align Images
             Messages{size(Messages,1)+1,1}=(sprintf('Aligning Image...\n'));
            app.OutputText.Value=Messages;
             
            % Find angle from vertical using estimates from both edges
            End_delxL=End_UL(1)-End_LL(1);
            End_delxR=End_UR(1)-End_LR(1);
            End_delyL=-(End_UL(2)-End_LL(2));
            End_delyR=-(End_UR(2)-End_LR(2));
            End_angleL=atand(End_delxL/End_delyL); %Positive is defined at a needing a CCW rotation to become vertical.
            End_angleR=atand(End_delxR/End_delyR);
            End_angle=mean([End_angleL,End_angleR]); %Avereage rotation needed

            End_width_U=pdist([End_UL;End_UR]);
            End_width_L=pdist([End_LL;End_LR]);
            End_height_L=pdist([End_UL;End_LL]);
            End_height_R=pdist([End_UR;End_LR]);

            Sheet_Width=mean([End_width_U, End_width_L]);
            Sheet_Height=mean([End_height_L, End_height_R]);

            %Find midpoint of film using mean of both diagonals
            End_middleULLR=[mean([End_UL(1),End_LR(1)]),mean([End_UL(2),End_LR(2)])];
            End_middleLLUR=[mean([End_LL(1),End_UR(1)]),mean([End_LL(2),End_UR(2)])];
            End_middle=mean([End_middleULLR;End_middleLLUR],1);

            %Find required shift to center film
            End_shift.x=size(I_end,2)-2*End_middle(1);
            End_shift.y=size(I_end,1)-2*End_middle(2);

            %Translate images
            I_end=imtranslate(I_end,[End_shift.x,End_shift.y],'OutputView','full','FillValues',255);

            %Rotate images
            I_end=imrotate(I_end,End_angle);

            %Crop down to only sheet
            minx=round(Sheet_Width-2*x_Crop);
            miny=round(Sheet_Height-2*y_Crop);

            if minx>size(I_end,2)
                Messages{size(Messages,1)+1,1}=(sprintf('WARNING!\nFilm width is bigger than the scanned images!\n'));
                app.OutputText.Value=Messages;
                 
            end
            if miny>size(I_end,1)
                Messages{size(Messages,1)+1,1}=(sprintf('WARNING!\nFilm width is bigger than the scanned images!\n'));
                app.OutputText.Value=Messages;
                 
            end

            I_end=imcrop(I_end,[(size(I_end,2)-minx)/2 (size(I_end,1)-miny)/2 minx miny]);

            %Resize Images
            I_end=imresize(I_end,Scale_Factor,'nearest');

            %ID circles in image
            h_circles=figure('Name','Identify Circles');
            set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
            end_mask=ones([size(I_end,1),size(I_end,2)]);
            imshow(I_end);

            for i=1:n_circles_end
                End_Circle_ID=CIRCLE_ID(h_circles);
                I_end_circle=I_end(max([1 round(End_Circle_ID(2)-CircleRad_max*Crop_Buffer)]):min(size(I_end,1),round(End_Circle_ID(2)+CircleRad_max*Crop_Buffer)),round(End_Circle_ID(1)-CircleRad_max*Crop_Buffer):round(End_Circle_ID(1)+CircleRad_max*Crop_Buffer),:); %Crop a small sqaure around the hole                [End_Circle_center, End_Circle_radii]=CIRCLE_HUNTER(I_end_circle,[round(CircleRad_min), round(CircleRad_max)],End_circle_color);
                End_Circle_center=End_Circle_center+[End_Circle_ID(1)-CircleRad_max*Crop_Buffer,End_Circle_ID(2)-CircleRad_max*Crop_Buffer]; %Put the coordinates back into the global coordinate system
                figure(h_circles); %bring figure to front
                End_Circle_center=round(End_Circle_center);
                End_Circle_radii=round(End_Circle_radii);
                viscircles(End_Circle_center,End_Circle_radii);
                I_end(End_Circle_center(2)-round(End_Circle_radii*Circle_NaN_buffer):End_Circle_center(2)+round(End_Circle_radii*Circle_NaN_buffer),End_Circle_center(1)-round(End_Circle_radii*Circle_NaN_buffer):End_Circle_center(1)+round(End_Circle_radii*Circle_NaN_buffer),1:3)=0; %mask out circle
                end_mask(End_Circle_center(2)-round(End_Circle_radii*Circle_NaN_buffer):End_Circle_center(2)+round(End_Circle_radii*Circle_NaN_buffer),End_Circle_center(1)-round(End_Circle_radii*Circle_NaN_buffer):End_Circle_center(1)+round(End_Circle_radii*Circle_NaN_buffer))=NaN; %mask out circle
            end

            %Set world coordinates
            World_end=imref2d(size(I_end),[-size(I_end,2)/2,size(I_end,2)/2],[-size(I_end,1)/2,size(I_end,1)/2]);

            %Show straightend images
            Messages{size(Messages,1)+1,1}=(sprintf('Alignment Complete\n'));
            app.OutputText.Value=Messages;
             
            h_aligned=figure('Name','Aligned Images');
            set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
            imshow(I_end);
            title('Ending Film');


            %% Save Images using the ending file name
            save(fullfile(End_path,erase(End_file,'.')),'I_end','end_mask','World_end');
            Messages{size(Messages,1)+1,1}=(sprintf('Saving image files as .mat\n'));
            app.OutputText.Value=Messages;
             
            
        case {'Fe','Mn'}     
            %% Image Import and Clean
            Start_fullfile=fullfile(Start_path,Start_file);
            End_fullfile=fullfile(End_path,End_file);
            
            I_start=imread(Start_fullfile);
            %Rotate images if upside down
            if Start_Upside_down_flag==1
                I_start=imrotate(I_start, 180);
            end

            I_end=imread(End_fullfile);
            if End_Upside_down_flag==1
                I_end=imrotate(I_end, 180);
            end

            %% Identify Corners in image
            fighan=figure;
            set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.

            %Starting Image - Corner selection
            subplot(1,2,1);
            imshow(I_start);
            User_Start_Corners=CORNER_ID(fighan);
            Start_UL=User_Start_Corners(1,:);
            Start_UR=User_Start_Corners(2,:);
            Start_LR=User_Start_Corners(3,:);
            Start_LL=User_Start_Corners(4,:);

            %End Image - Corner selection
            subplot(1,2,2);
            imshow(I_end);
            User_End_Corners=CORNER_ID(fighan);
            End_UL=User_End_Corners(1,:);
            End_UR=User_End_Corners(2,:);
            End_LR=User_End_Corners(3,:);
            End_LL=User_End_Corners(4,:);

            close(fighan);

             %% Align Images
             Messages{size(Messages,1)+1,1}=(sprintf('Aligning Images...\n'));
            app.OutputText.Value=Messages;
             
            % Find angle from vertical using estimates from both edges
            Start_delxL=Start_UL(1)-Start_LL(1);
            Start_delxR=Start_UR(1)-Start_LR(1);
            Start_delyL=-(Start_UL(2)-Start_LL(2));
            Start_delyR=-(Start_UR(2)-Start_LR(2));
            Start_angleL=atand(Start_delxL/Start_delyL); %Positive is defined at a needing a CCW rotation to become vertical.
            Start_angleR=atand(Start_delxR/Start_delyR);
            Start_angle=mean([Start_angleL,Start_angleR]); %Avereage rotation needed

            End_delxL=End_UL(1)-End_LL(1);
            End_delxR=End_UR(1)-End_LR(1);
            End_delyL=-(End_UL(2)-End_LL(2));
            End_delyR=-(End_UR(2)-End_LR(2));
            End_angleL=atand(End_delxL/End_delyL); %Positive is defined at a needing a CCW rotation to become vertical.
            End_angleR=atand(End_delxR/End_delyR);
            End_angle=mean([End_angleL,End_angleR]); %Avereage rotation needed

            % Find the size of the sheet
            Start_width_U=pdist([Start_UL;Start_UR]);
            Start_width_L=pdist([Start_LL;Start_LR]);
            Start_height_L=pdist([Start_UL;Start_LL]);
            Start_height_R=pdist([Start_UR;Start_LR]);

            End_width_U=pdist([End_UL;End_UR]);
            End_width_L=pdist([End_LL;End_LR]);
            End_height_L=pdist([End_UL;End_LL]);
            End_height_R=pdist([End_UR;End_LR]);

            Sheet_Width=mean([Start_width_U, Start_width_L,End_width_U, End_width_L]);
            Sheet_Height=mean([Start_height_L, Start_height_R,End_height_L, End_height_R]);

            %Find midpoint of film using mean of both diagonals
            Start_middleULLR=[mean([Start_UL(1),Start_LR(1)]),mean([Start_UL(2),Start_LR(2)])];
            Start_middleLLUR=[mean([Start_LL(1),Start_UR(1)]),mean([Start_LL(2),Start_UR(2)])];
            Start_middle=mean([Start_middleULLR;Start_middleLLUR],1);

            End_middleULLR=[mean([End_UL(1),End_LR(1)]),mean([End_UL(2),End_LR(2)])];
            End_middleLLUR=[mean([End_LL(1),End_UR(1)]),mean([End_LL(2),End_UR(2)])];
            End_middle=mean([End_middleULLR;End_middleLLUR],1);

            %Find required shift to center film
            Start_shift.x=size(I_start,2)-2*Start_middle(1);
            Start_shift.y=size(I_start,1)-2*Start_middle(2);
            End_shift.x=size(I_end,2)-2*End_middle(1);
            End_shift.y=size(I_end,1)-2*End_middle(2);

            %Translate images
            I_start=imtranslate(I_start,[Start_shift.x,Start_shift.y],'OutputView','full','FillValues',0);
            I_end=imtranslate(I_end,[End_shift.x,End_shift.y],'OutputView','full','FillValues',255);

            %Rotate images
            I_start=imrotate(I_start,Start_angle);
            I_end=imrotate(I_end,End_angle);

            %Crop down to only sheet
            minx=round(Sheet_Width-2*x_Crop);
            miny=round(Sheet_Height-2*y_Crop);

            if minx>size(I_start,2) || minx>size(I_end,2)
                Messages{size(Messages,1)+1,1}=(sprintf('WARNING!\nFilm width is bigger than one of the scanned images!\n'));
                app.OutputText.Value=Messages;
            end
            if miny>size(I_start,1) || miny>size(I_end,1)
                Messages{size(Messages,1)+1,1}=(sprintf('WARNING!\nFilm width is bigger than one of the scanned images!\n'));
                app.OutputText.Value=Messages; 
            end

            I_start=imcrop(I_start,[(size(I_start,2)-minx)/2 (size(I_start,1)-miny)/2 minx miny]);
            I_end=imcrop(I_end,[(size(I_end,2)-minx)/2 (size(I_end,1)-miny)/2 minx miny]);

            %Resize Images
            I_start=imresize(I_start,Scale_Factor,'nearest');
            I_end=imresize(I_end,Scale_Factor,'nearest');

            %ID circles in starting image
            start_mask=ones([size(I_start,1),size(I_start,2)]);
            h_circles=figure('Name','Identify Circles');
            set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
            subplot(1,2,1);
            imshow(I_start);

            for i=1:n_circles_start
                Start_Circle_ID=CIRCLE_ID(h_circles);
                I_start_circle=I_start(max([1 round(Start_Circle_ID(2)-CircleRad_max*Crop_Buffer)]):min(size(I_start,1),round(Start_Circle_ID(2)+CircleRad_max*Crop_Buffer)),round(Start_Circle_ID(1)-CircleRad_max*Crop_Buffer):round(Start_Circle_ID(1)+CircleRad_max*Crop_Buffer),:); %Crop a small sqaure around the hole
                [Start_Circle_center, Start_Circle_radii]=CIRCLE_HUNTER(I_start_circle,[round(CircleRad_min), round(CircleRad_max)],Start_circle_color);
                Start_Circle_center=Start_Circle_center+[Start_Circle_ID(1)-CircleRad_max*Crop_Buffer,Start_Circle_ID(2)-CircleRad_max*Crop_Buffer]; %Put the coordinates back into the global coordinate system
                figure(h_circles); %bring figure to front
                Start_Circle_center=round(Start_Circle_center);
                Start_Circle_radii=round(Start_Circle_radii);
                viscircles(Start_Circle_center,Start_Circle_radii);
                I_start(Start_Circle_center(2)-round(Start_Circle_radii*Circle_NaN_buffer):Start_Circle_center(2)+round(Start_Circle_radii*Circle_NaN_buffer),Start_Circle_center(1)-round(Start_Circle_radii*Circle_NaN_buffer):Start_Circle_center(1)+round(Start_Circle_radii*Circle_NaN_buffer),1:3)=0; %mask out circle
                start_mask(Start_Circle_center(2)-round(Start_Circle_radii*Circle_NaN_buffer):Start_Circle_center(2)+round(Start_Circle_radii*Circle_NaN_buffer),Start_Circle_center(1)-round(Start_Circle_radii*Circle_NaN_buffer):Start_Circle_center(1)+round(Start_Circle_radii*Circle_NaN_buffer))=NaN; %mask out circle
            end

            %ID circles in ending image
            end_mask=ones([size(I_end,1),size(I_end,2)]);
            subplot(1,2,2);
            imshow(I_end);

            for i=1:n_circles_end
                End_Circle_ID=CIRCLE_ID(h_circles);
                I_end_circle=I_end(max([1 round(End_Circle_ID(2)-CircleRad_max*Crop_Buffer)]):min(size(I_end,1),round(End_Circle_ID(2)+CircleRad_max*Crop_Buffer)),round(End_Circle_ID(1)-CircleRad_max*Crop_Buffer):round(End_Circle_ID(1)+CircleRad_max*Crop_Buffer),:); %Crop a small sqaure around the hole
                [End_Circle_center, End_Circle_radii]=CIRCLE_HUNTER(I_end_circle,[round(CircleRad_min), round(CircleRad_max)],End_circle_color);
                End_Circle_center=End_Circle_center+[End_Circle_ID(1)-CircleRad_max*Crop_Buffer,End_Circle_ID(2)-CircleRad_max*Crop_Buffer]; %Put the coordinates back into the global coordinate system
                figure(h_circles); %bring figure to front
                End_Circle_center=round(End_Circle_center);
                End_Circle_radii=round(End_Circle_radii);
                viscircles(End_Circle_center,End_Circle_radii);
                I_end(End_Circle_center(2)-round(End_Circle_radii*Circle_NaN_buffer):End_Circle_center(2)+round(End_Circle_radii*Circle_NaN_buffer),End_Circle_center(1)-round(End_Circle_radii*Circle_NaN_buffer):End_Circle_center(1)+round(End_Circle_radii*Circle_NaN_buffer),1:3)=0; %mask out circle
                end_mask(End_Circle_center(2)-round(End_Circle_radii*Circle_NaN_buffer):End_Circle_center(2)+round(End_Circle_radii*Circle_NaN_buffer),End_Circle_center(1)-round(End_Circle_radii*Circle_NaN_buffer):End_Circle_center(1)+round(End_Circle_radii*Circle_NaN_buffer))=NaN; %mask out circle
            end

            %Set world coordinates
            World_start=imref2d(size(I_start),[-size(I_start,2)/2,size(I_start,2)/2],[-size(I_start,1)/2,size(I_start,1)/2]);
            World_end=imref2d(size(I_end),[-size(I_end,2)/2,size(I_end,2)/2],[-size(I_end,1)/2,size(I_end,1)/2]);

            %Show straightend images
            Messages{size(Messages,1)+1,1}=(sprintf('Alignment Complete\n'));
            app.OutputText.Value=Messages;
             
            h_aligned=figure('Name','Aligned Images');
            set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
            subplot(1,4,1);
            imshow(I_start);

            title('Starting Film');
            subplot(1,4,2);
            imshow(I_end);
            title('Ending Film');
            subplot(1,4,3);
            imshowpair(I_start, World_start, I_end, World_end,'checkerboard');
            title('Overlapped Films');
            subplot(1,4,4);
            imshowpair(I_start, World_start, I_end, World_end,'diff');
            title('Difference between films');

            %% Save Images using the starting file name
            save(fullfile(End_path,erase(End_file,'.')),'I_start','I_end','start_mask','end_mask','World_start','World_end');
            Messages{size(Messages,1)+1,1}=(sprintf('Saving image files as .mat\n'));
            app.OutputText.Value=Messages;
             
        otherwise
            Messages{size(Messages,1)+1,1}=(sprintf('Improper IRIS film type specified in "AdjustableParameters" spreadsheet'));
            app.OutputText.Value=Messages;
             
            return;
    end
end

switch IRIS_Type
    case {'Mn','Fe'}
        %% Convert Colors to CIELAB
        I_start_lab=rgb2lab(I_start,'ColorSpace','srgb');
        I_start_L=I_start_lab(:,:,1).*start_mask;
        I_start_a=I_start_lab(:,:,2).*start_mask;
        I_start_b=I_start_lab(:,:,3).*start_mask;
        I_end_lab=rgb2lab(I_end,'ColorSpace','srgb');
        I_end_L=I_end_lab(:,:,1).*end_mask;
        I_end_a=I_end_lab(:,:,2).*end_mask;
        I_end_b=I_end_lab(:,:,3).*end_mask;

        lab=figure('Name','CIELAB Difference Images');
        set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
        subplot(1,3,1);
        imshowpair(I_start_L, World_start, I_end_L, World_end,'diff');
        title('Lightness');
        subplot(1,3,2);
        imshowpair(I_start_a, World_start, I_end_a, World_end,'diff');
        title('a');
        subplot(1,3,3);
        imshowpair(I_start_b, World_start, I_end_b, World_end,'diff');
        title('b');

        %% Cluster Analysis - Starting Image
        I_start_lab_single=im2single(I_start_lab);
        [I_start_labels, I_start_clustercolors]=imsegkmeans(I_start_lab_single,n_clusters_start,'NormalizeInput',false); %Do not rescale channels
        I_start_clustercolors_rgb=lab2rgb(I_start_clustercolors,'OutputType','double'); %Convert colors to rgb
        I_start_clustercolors_rgb(I_start_clustercolors_rgb<0)=0; %Eliminate negative values
        I_start_clustercolors_rgb(I_start_clustercolors_rgb>1)=1; %Eliminate values above unity
        I_start_clustered=label2rgb(I_start_labels,I_start_clustercolors_rgb);
        I_start_overlay = labeloverlay(I_start,I_start_labels,'Transparency',0);
        I_start_overlay_color = labeloverlay(I_start,I_start_labels,'Transparency',0,'Colormap',I_start_clustercolors_rgb);
        figure('Name','Clustered Images');
        set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
        subplot(1,4,1);
        imshow(I_start_overlay);
        title('Starting Image')
        subplot(1,4,2);
        imshow(I_start_overlay_color);
        title('Starting Image')

        I_start_Munsell=MUNSELL_MATCH(I_start_clustercolors);

        %Calculate area of each cluster
        mask=zeros([size(I_start_labels,1),size(I_start_labels,2),n_clusters_start]);
        Start_cluster_area=zeros(1,n_clusters_start);
        for i=1:n_clusters_start
            mask(:,:,i)=I_start_labels==i;
            Start_cluster_area(i)=sum(sum(mask(:,:,i)))./numel(mask(:,:,1))*100;
        end

        %% Cluster Analysis - Ending Image
        I_end_lab_single=im2single(I_end_lab);
        [I_end_labels, I_end_clustercolors]=imsegkmeans(I_end_lab_single,n_clusters_end,'NormalizeInput',false); %Do not rescale channels
        I_end_clustercolors_rgb=lab2rgb(I_end_clustercolors,'OutputType','double'); %Convert colors to rgb
        I_end_clustercolors_rgb(I_end_clustercolors_rgb<0)=0; %Eliminate negative values
        I_end_clustercolors_rgb(I_end_clustercolors_rgb>1)=1; %Eliminate values above unity
        I_end_clustered=label2rgb(I_end_labels,I_end_clustercolors_rgb);
        I_end_overlay = labeloverlay(I_end,I_end_labels,'Transparency',0);
        I_end_overlay_color = labeloverlay(I_end,I_end_labels,'Transparency',0,'Colormap',I_end_clustercolors_rgb);
        subplot(1,4,3);
        imshow(I_end_overlay);
        title('Ending Image');
        subplot(1,4,4);
        imshow(I_end_overlay_color);
        title('Ending Image');

        I_end_Munsell=MUNSELL_MATCH(I_end_clustercolors);

        %Calculate area of each cluster
        mask=zeros([size(I_end_labels,1),size(I_end_labels,2),n_clusters_end]);
        End_cluster_area=zeros(1,n_clusters_end);
        for i=1:n_clusters_end
            mask(:,:,i)=I_end_labels==i;
            End_cluster_area(i)=sum(sum(mask(:,:,i)))./numel(mask(:,:,1))*100;
        end


        %% Quantify Intensity Change
        I_L_diff=(I_end_L-I_start_L)./(White_L-I_start_L)*100; %Convert to a percent removal
        I_L_diff(I_L_diff<0)=0; %Limit removal to be >0 for each pixel
        I_L_diff(I_L_diff>100)=100; %Limit removal to be <100% for each pixel

        figure('Name','Percent Removal');
        imshow(I_L_diff,'DisplayRange',[0 100]); colorbar;
        set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.

        I_L_diff(I_L_diff==Inf|I_L_diff==-Inf)=NaN; %Change infs to NaNs
        mean_I_L_diff_y=nanmean(I_L_diff,2);
        std_I_L_diff_y=nanstd(I_L_diff,0,2);
        mean_I_L_diff_x=nanmean(I_L_diff,1);
        std_I_L_diff_x=nanstd(I_L_diff,0,1);
        median_I_L_diff=nanmedian(I_L_diff,'all');
        L_trend=figure('Name','Trends in removal');
        subplot(1,2,1)
        x_width=1:length(mean_I_L_diff_x);
        x_width=x_width/DPI/Scale_Factor; %convert to inches
        plot(x_width,mean_I_L_diff_x,'-b')
        hold on;
        plot(x_width,mean_I_L_diff_x+std_I_L_diff_x,':b')
        plot(x_width,mean_I_L_diff_x-std_I_L_diff_x,':b')
        axis([0 inf 0 100]);
        xlabel('Sheet Width (in)');
        ylabel('Average % Removal')
        subplot(1,2,2)
        y_length=1:length(mean_I_L_diff_y);
        y_length=y_length/DPI/Scale_Factor;
        plot(y_length,mean_I_L_diff_y,'-r')
        hold on;
        plot(y_length,mean_I_L_diff_y+std_I_L_diff_y,':r')
        plot(y_length,mean_I_L_diff_y-std_I_L_diff_y,':r')
        axis([0 inf 0 100]);
        ylabel('Average % Removal')
        xlabel('Sheet Depth (in)');

        %% Fit Lines to Lightness Data
        Messages{size(Messages,1)+1,1}=(sprintf('Fitting trendlines...\n'));
        app.OutputText.Value=Messages;
        xx=y_length';
        yy=mean_I_L_diff_y;
        break_index=findchangepts(yy); %Find breakpoint
        StartPoints=[1, 50, 0, 100, xx(break_index)]; %Starting guesses for function
        Lower=[-inf, -inf, -inf, -inf, 0]; %Lower limits
        Upper=[inf, inf, inf, inf, y_length(length(y_length))];
        ft=fittype( 'PIECEWISELINE( x, a, b, c, d, e)');
        f=fit(xx, yy, ft, 'StartPoint', StartPoints,'Lower',Lower,'Upper',Upper);
        plot( xx, PIECEWISELINE(xx,f.a,f.b,f.c,f.d,f.e),'k--');

        %% Quantify Hue Change
        I_a_diff=(I_end_a-I_start_a);
        I_b_diff=(I_end_b-I_start_b);

        figure('Name','Change in Film Color');
        set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
        subplot(1,3,1)
        CIELAB_PLOT([nanmean(nanmean(I_start_a)),nanmean(nanmean(I_start_b))],[nanmean(nanmean(I_end_a)),nanmean(nanmean(I_end_b))]);
        subplot(1,3,2);

        I_a_diff_color=I_start_a;
        I_a_diff_color_r=max(0,I_a_diff)./50;
        I_a_diff_color_g=max(0,-I_a_diff)./50;
        I_a_diff_color_b=zeros(size(I_a_diff_color_g));
        I_a_diff_color_rgb=I_start_a;
        I_a_diff_color_rgb(:,:,1)=I_a_diff_color_r;
        I_a_diff_color_rgb(:,:,2)=I_a_diff_color_g;
        I_a_diff_color_rgb(:,:,3)=I_a_diff_color_b;
        imshow(I_a_diff_color_rgb);
        title('Change in a* from start to end')
        subplot(1,3,3);
        I_b_diff_color=I_start_b;
        I_b_diff_color_r=max(0,I_b_diff)./50;
        I_b_diff_color_g=max(0,I_b_diff)./50;
        I_b_diff_color_b=max(0,-I_b_diff)./50;
        I_b_diff_color_rgb=I_start_b;
        I_b_diff_color_rgb(:,:,1)=I_b_diff_color_r;
        I_b_diff_color_rgb(:,:,2)=I_b_diff_color_g;
        I_b_diff_color_rgb(:,:,3)=I_b_diff_color_b;
        imshow(I_b_diff_color_rgb);
        title('Change in b* from start to end')

    case 'S'
        %% Convert Colors to CIELAB
        I_end_lab=rgb2lab(I_end,'ColorSpace','srgb');
        I_end_L=I_end_lab(:,:,1).*end_mask;
        I_end_a=I_end_lab(:,:,2).*end_mask;
        I_end_b=I_end_lab(:,:,3).*end_mask;

        %% Cluster Analysis - Ending Image
        I_end_lab_single=im2single(I_end_lab);
        [I_end_labels, I_end_clustercolors]=imsegkmeans(I_end_lab_single,n_clusters_end,'NormalizeInput',false); %Do not rescale channels
        I_end_clustercolors_rgb=lab2rgb(I_end_clustercolors,'OutputType','double'); %Convert colors to rgb
        I_end_clustercolors_rgb(I_end_clustercolors_rgb<0)=0; %Eliminate negative values
        I_end_clustercolors_rgb(I_end_clustercolors_rgb>1)=1; %Eliminate values above unity
        I_end_clustered=label2rgb(I_end_labels,I_end_clustercolors_rgb);
        I_end_overlay = labeloverlay(I_end,I_end_labels,'Transparency',0);
        I_end_overlay_color = labeloverlay(I_end,I_end_labels,'Transparency',0,'Colormap',I_end_clustercolors_rgb);
        figure('Name','Clustered Images');
        set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
        subplot(1,2,1);
        imshow(I_end_overlay);
        title('Ending Image');
        subplot(1,2,2);
        imshow(I_end_overlay_color);
        title('Ending Image');

        I_end_Munsell=MUNSELL_MATCH(I_end_clustercolors);

        %Calculate area of each cluster
        mask=zeros([size(I_end_labels,1),size(I_end_labels,2),n_clusters_end]);
        End_cluster_area=zeros(1,n_clusters_end);
        for i=1:n_clusters_end
            mask(:,:,i)=I_end_labels==i;
            End_cluster_area(i)=sum(sum(mask(:,:,i)))./numel(mask(:,:,1))*100;
        end

        %%Masking out Fe
        Fe_Mask=I_end_b>Fe_b_Threshold; %Pixels where b* is more than threshold can represent Fe
        S_Percent=sum(sum(~Fe_Mask))/numel(Fe_Mask)*100; %The percentage of pixels that are S
        I_sulfide=100-I_end_L; %Estimate sulfide using lightness
        I_sulfide(Fe_Mask)=0; %Remove Fe pixels
        

        figure('Name','Sulfide Plots');
        set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
        subplot(1,3,1)
        imshow(I_end);
        title('Ending Film')
        subplot(1,3,2);
        imshow(~Fe_Mask,'DisplayRange',[0 1]); colorbar;
        title('S Pixels')
        subplot(1,3,3);
        imshow(I_sulfide,'DisplayRange',[0 100]); colorbar;
        title('Relative intensity of sulfide')

        %Calculate statistics for films
        mean_S_y=nanmean(I_sulfide,2);
        std_S_y=nanstd(I_sulfide,0,2);
        mean_S_x=nanmean(I_sulfide,1);
        std_S_x=nanstd(I_sulfide,0,1);
        median_S=nanmedian(I_sulfide,'all');
        
        x_width=1:length(mean_S_x);
        x_width=x_width/DPI/Scale_Factor; %convert to inches
        y_length=1:length(mean_S_y);
        y_length=y_length/DPI/Scale_Factor;

        figure('Name','Trends in relative intensity of sulfide');
        subplot(1,2,1)
        plot(x_width,mean_S_x,'-b')
        hold on;
        plot(x_width,mean_S_x+std_S_x,':b')
        plot(x_width,mean_S_x-std_S_x,':b')
        axis([0 inf 0 100]);
        xlabel('Sheet Width (in)');
        ylabel('Relative intensity of sulfide across film')
        subplot(1,2,2)
        plot(y_length,mean_S_y,'-r')
        hold on;
        plot(y_length,mean_S_y+std_S_y,':r')
        plot(y_length,mean_S_y-std_S_y,':r')
        axis([0 inf 0 100]);
        ylabel('Relative intensity of sulfide with depth')
        xlabel('Sheet Depth (in)');

        % Fit Lines to Lightness Data
        Messages{size(Messages,1)+1,1}=(sprintf('Fitting trendlines...\n'));
        app.OutputText.Value=Messages;
        xx=y_length';
        yy_S=mean_S_y;
        break_index=findchangepts(yy_S); %Find breakpoint
        StartPoints=[1, 50, 0, 100, xx(break_index)]; %Starting guesses for function
        Lower=[-inf, -inf, -inf, -inf, 0]; %Lower limits
        Upper=[inf, inf, inf, inf, y_length(length(y_length))];
        ft=fittype( 'PIECEWISELINE( x, a, b, c, d, e)');
        f_S=fit(xx, yy_S, ft, 'StartPoint', StartPoints,'Lower',Lower,'Upper',Upper);
        plot( xx, PIECEWISELINE(xx,f_S.a,f_S.b,f_S.c,f_S.d,f_S.e),'k--');
        
end

%% Masking out Fe on Mn Films
switch IRIS_Type
    case 'Mn'
        Fe_Mask=I_b_diff>Fe_b_Threshold; %Pixels where b* changes by more than threshold can represent Fe
        Fe_Percent=sum(sum(Fe_Mask))/numel(Fe_Mask)*100; %The percentage of pixels that are Fe
        I_L_diff_noFe=I_L_diff; %create variable for films ignoring Fe pixels
        I_L_diff_noFe(Fe_Mask==1)=100; %count Fe pixels as 100% removal

        figure('Name','Fe pixels (for use with Mn films)');
        set(gcf, 'Position', get(0,'Screensize')); % Maximize figure.
        subplot(1,3,1)
        imshow(I_end);
        title('Ending Film')
        subplot(1,3,2);
        imshow(Fe_Mask,'DisplayRange',[0 1]); colorbar;
        title('Fe Pixels')
        subplot(1,3,3);
        imshow(I_L_diff_noFe,'DisplayRange',[0 100]); colorbar;
        title('Percent removal considering Fe pixels as 100% removal')

        %Calculate statistics for films when considering Fe pixels as 100% removal
        mean_I_L_diff_y_noFe=nanmean(I_L_diff_noFe,2);
        std_I_L_diff_y_noFe=nanstd(I_L_diff_noFe,0,2);
        mean_I_L_diff_x_noFe=nanmean(I_L_diff_noFe,1);
        std_I_L_diff_x_noFe=nanstd(I_L_diff_noFe,0,1);
        median_I_L_diff_noFe=nanmedian(I_L_diff_noFe,'all');

        figure('Name','Trends in removal considering Fe pixels as 100% removed');
        subplot(1,2,1)
        plot(x_width,mean_I_L_diff_x_noFe,'-b')
        hold on;
        plot(x_width,mean_I_L_diff_x_noFe+std_I_L_diff_x_noFe,':b')
        plot(x_width,mean_I_L_diff_x_noFe-std_I_L_diff_x_noFe,':b')
        axis([0 inf 0 100]);
        xlabel('Sheet Width (in)');
        ylabel('Average % Removal (Fe pixels =100% removal)')
        subplot(1,2,2)
        plot(y_length,mean_I_L_diff_y_noFe,'-r')
        hold on;
        plot(y_length,mean_I_L_diff_y_noFe+std_I_L_diff_y_noFe,':r')
        plot(y_length,mean_I_L_diff_y_noFe-std_I_L_diff_y_noFe,':r')
        axis([0 inf 0 100]);
        ylabel('Average % Removal (Fe pixels =100% removal)')
        xlabel('Sheet Depth (in)');

        % Fit Lines to Lightness Data
        Messages{size(Messages,1)+1,1}=(sprintf('Fitting trendlines...\n'));
        app.OutputText.Value=Messages;
        yy_noFe=mean_I_L_diff_y_noFe;
        break_index=findchangepts(yy_noFe); %Find breakpoint
        StartPoints=[1, 50, 0, 100, xx(break_index)]; %Starting guesses for function
        f_noFe=fit(xx, yy_noFe, ft, 'StartPoint', StartPoints,'Lower',Lower,'Upper',Upper);
        plot( xx, PIECEWISELINE(xx,f_noFe.a,f_noFe.b,f_noFe.c,f_noFe.d,f_noFe.e),'k--');
end


%% Output Data
Messages{size(Messages,1)+1,1}=(sprintf('Writing output data to Excel...\n'));
    app.OutputText.Value=Messages;

    OutFile=fullfile(End_path,OutFile); %Full file path - output file with ending images
if isfile(OutFile)
    existing_output=readtable(OutFile,'Sheet',OutSheet);
    existing_rows=1+size(existing_output,1);
else
    existing_rows=1;
    Messages{size(Messages,1)+1,1}=(sprintf('Output file does not exist. Writing headers...\n'));
    app.OutputText.Value=Messages;
     
    switch IRIS_Type %Adjust headers for each element
        case 'Mn'
            Headers={'Start File Name','End File Name','Sheet Width (in)','Sheet Length (in)','Mean % Removal','Median % Removal'...
                'Mean L value of ending film','Mean vertical standard deviation of removal','Mean horizontal standard deviation of removal',...
                'Overall standard deviation of removal','Removal centroid in x','Removal centroid in y',...
                'Percent of Film with Fe pixels (for Mn films)', 'Mean % Removal assuming Fe pixels as 100% removal (for Mn Films)','Median % Removal assuming Fe pixels as 100% removal (for Mn Films)'...
                'Mean vertical standard deviation of removal assuming Fe pixels as 100% removal (for Mn Films)','Mean horizontal standard deviation of removal assuming Fe pixels as 100% removal (for Mn Films)',...
                'Overall standard deviation of removal assuming Fe pixels as 100% removal (for Mn Films)','Removal centroid in x assuming Fe pixels as 100% removal (for Mn Films)','Removal centroid in y assuming Fe pixels as 100% removal (for Mn Films)',...
                'Film change in a*', 'Film change in b*', 'Starting film cluster areas (%)','','','Starting film Munsell cluster colors','','',...
                'Ending film cluster areas (%)','','','','','Ending film Munsell cluster colors','','','','',...
                'Removal fxn Intercept1','Removal fxn Slope1','Removal fxn Intercept2','Removal fxn Slope2','Removal fxn Breakpoint (in)',...
                'Removal fxn Intercept1 (no Fe)','Removal fxn Slope1 (no Fe)','Removal fxn Intercept2 (no Fe)','Removal fxn Slope2 (no Fe)','Removal fxn Breakpoint (no Fe) (in)'};
        case 'Fe'
             Headers={'Start File Name','End File Name','Sheet Width (in)','Sheet Length (in)','Mean % Removal','Median % Removal'...
            'Mean L value of ending film','Mean vertical standard deviation of removal','Mean horizontal standard deviation of removal',...
            'Overall standard deviation of removal','Removal centroid in x','Removal centroid in y',...          
            'Film change in a*', 'Film change in b*', 'Starting film cluster areas (%)','','','Starting film Munsell cluster colors','','',...
            'Ending film cluster areas (%)','','','','','Ending film Munsell cluster colors','','','','',...
            'Removal fxn Intercept1','Removal fxn Slope1','Removal fxn Intercept2','Removal fxn Slope2','Removal fxn Breakpoint (in)'};
        case 'S'
             Headers={'End File Name','Sheet Width (in)','Sheet Length (in)','Mean Relative S','Median Relative S','Percent of S Pixels'...
            'Mean L value of ending film','Mean vertical standard deviation of S','Mean horizontal standard deviation of S',...
            'Overall standard deviation of S','S centroid in x','S centroid in y',...          
            'Film avg a*', 'Film avg b*',...
            'Ending film cluster areas (%)','','','','','Ending film Munsell cluster colors','','','','',...
            'S fxn Intercept1','S fxn Slope1','S fxn Intercept2','S fxn Slope2','S fxn Breakpoint (in)'};
    end
    writecell(Headers,OutFile,'Sheet',OutSheet);
end

switch IRIS_Type %Adjust output data based upon film type
    case 'Mn'
        mean_I_L_val=nanmean(nanmean(I_end_L)); %mean values of L

        mean_I_L_diff=nanmean(mean_I_L_diff_y); %mean change in saturation
        mean_std_I_L_diff_x=nanmean(std_I_L_diff_x); %mean standard deviation moving across x-direction for removal (i.e., vertical std)
        mean_std_I_L_diff_y=nanmean(std_I_L_diff_y); %mean standard deviation moving across y-direction for removal (i.d., horizontal std)
        mean_std_I_L_diff=nanstd(reshape(I_L_diff,[],1)); %standard deviation for entire sheet for removal

        x_cent_I_L_diff=sum(x_width.*mean_I_L_diff_x)./sum(mean_I_L_diff_x)-median(x_width); %Location of removal centroid in x-direction (0 is center of sheet, positive right)
        y_cent_I_L_diff=sum(y_length'.*mean_I_L_diff_y)./sum(mean_I_L_diff_y)-median(y_length); %Location of removal centroid in y-direction (0 is center of sheet, positive down)

        % Values assuming Fe pixels are 100% removal
        mean_I_L_diff_noFe=nanmean(mean_I_L_diff_y_noFe); %mean change in saturation
        mean_std_I_L_diff_x_noFe=nanmean(std_I_L_diff_x_noFe); %mean standard deviation moving across x-direction for removal (i.e., vertical std)
        mean_std_I_L_diff_y_noFe=nanmean(std_I_L_diff_y_noFe); %mean standard deviation moving across y-direction for removal (i.d., horizontal std)
        mean_std_I_L_diff_noFe=nanstd(reshape(I_L_diff_noFe,[],1)); %standard deviation for entire sheet for removal

        x_cent_I_L_diff_noFe=sum(x_width.*mean_I_L_diff_x_noFe)./sum(mean_I_L_diff_x_noFe)-median(x_width); %Location of removal centroid in x-direction (0 is center of sheet, positive right)
        y_cent_I_L_diff_noFe=sum(y_length'.*mean_I_L_diff_y_noFe)./sum(mean_I_L_diff_y_noFe)-median(y_length); %Location of removal centroid in y-direction (0 is center of sheet, positive down)

        mean_I_a_diff=nanmean(nanmean(I_a_diff)); %mean change in a*
        mean_I_b_diff=nanmean(nanmean(I_b_diff)); %mean change in b*

        % Output sheet width/length, mean % removal, mean std in % removal for x &
        % y, overall std in % removal, x & y centroids for % removal, mean change
        % in a* and b*, starting cluster area, starting cluster colors, ending
        % cluster area, ending cluster colors, Removal fit values
        Outdata=table({Start_file}, {End_file},x_width(length(x_width)), y_length(length(y_length)),...
            mean_I_L_diff, median_I_L_diff,...
            mean_I_L_val, mean_std_I_L_diff_x, mean_std_I_L_diff_y,...
            mean_std_I_L_diff, x_cent_I_L_diff, y_cent_I_L_diff,...
            Fe_Percent, mean_I_L_diff_noFe, median_I_L_diff_noFe,...
            mean_std_I_L_diff_x_noFe, mean_std_I_L_diff_y_noFe,...
            mean_std_I_L_diff_noFe, x_cent_I_L_diff_noFe, y_cent_I_L_diff_noFe,...
            mean_I_a_diff, mean_I_b_diff,...
            Start_cluster_area, I_start_Munsell', num2cell(End_cluster_area), I_end_Munsell',...
            f.a,f.b,f.c,f.d,f.e,f_noFe.a,f_noFe.b,f_noFe.c,f_noFe.d,f_noFe.e);
        
    case 'Fe'
        mean_I_L_val=nanmean(nanmean(I_end_L)); %mean values of L

        mean_I_L_diff=nanmean(mean_I_L_diff_y); %mean change in saturation
        mean_std_I_L_diff_x=nanmean(std_I_L_diff_x); %mean standard deviation moving across x-direction for removal (i.e., vertical std)
        mean_std_I_L_diff_y=nanmean(std_I_L_diff_y); %mean standard deviation moving across y-direction for removal (i.d., horizontal std)
        mean_std_I_L_diff=nanstd(reshape(I_L_diff,[],1)); %standard deviation for entire sheet for removal

        x_cent_I_L_diff=sum(x_width.*mean_I_L_diff_x)./sum(mean_I_L_diff_x)-median(x_width); %Location of removal centroid in x-direction (0 is center of sheet, positive right)
        y_cent_I_L_diff=sum(y_length'.*mean_I_L_diff_y)./sum(mean_I_L_diff_y)-median(y_length); %Location of removal centroid in y-direction (0 is center of sheet, positive down)

        mean_I_a_diff=nanmean(nanmean(I_a_diff)); %mean change in a*
        mean_I_b_diff=nanmean(nanmean(I_b_diff)); %mean change in b*

        % Output sheet width/length, mean % removal, mean std in % removal for x &
        % y, overall std in % removal, x & y centroids for % removal, mean change
        % in a* and b*, starting cluster area, starting cluster colors, ending
        % cluster area, ending cluster colors, Removal fit values
        Outdata=table({Start_file}, {End_file}, x_width(length(x_width)), y_length(length(y_length)),...
            mean_I_L_diff, median_I_L_diff,...
            mean_I_L_val, mean_std_I_L_diff_x, mean_std_I_L_diff_y,...
            mean_std_I_L_diff, x_cent_I_L_diff, y_cent_I_L_diff,...
            mean_I_a_diff, mean_I_b_diff,...
            Start_cluster_area, I_start_Munsell', num2cell(End_cluster_area), I_end_Munsell',...
            f.a,f.b,f.c,f.d,f.e);
        
    case 'S'
        mean_I_L_val=nanmean(nanmean(I_end_L)); %mean values of L

        mean_S=nanmean(mean_S_y); %mean sulfide
        mean_std_S_x=nanmean(std_S_x); %mean standard deviation moving across x-direction for S (i.e., vertical std)
        mean_std_S_y=nanmean(std_S_y); %mean standard deviation moving across y-direction for S (i.d., horizontal std)
        mean_std_S=nanstd(reshape(I_sulfide,[],1)); %standard deviation for entire sheet for S

        x_cent_S=sum(x_width.*mean_S_x)./sum(mean_S_x)-median(x_width); %Location of removal centroid in x-direction (0 is center of sheet, positive right)
        y_cent_S=sum(y_length'.*mean_S_y)./sum(mean_S_y)-median(y_length); %Location of removal centroid in y-direction (0 is center of sheet, positive down)

        mean_I_a=nanmean(nanmean(I_end_a)); %mean a*
        mean_I_b=nanmean(nanmean(I_end_b)); %mean b*

        % Output sheet width/length, mean % removal, mean std in % removal for x &
        % y, overall std in % removal, x & y centroids for % removal, mean change
        % in a* and b*, starting cluster area, starting cluster colors, ending
        % cluster area, ending cluster colors, Removal fit values
        Outdata=table({End_file}, x_width(length(x_width)), y_length(length(y_length)),...
            mean_S, median_S, S_Percent,...
            mean_I_L_val, mean_std_S_x, mean_std_S_y,...
            mean_std_S, x_cent_S, y_cent_S,...
            mean_I_a, mean_I_b, End_cluster_area, I_end_Munsell',...
            f_S.a,f_S.b,f_S.c,f_S.d,f_S.e);     
end

%Write file
writetable(Outdata,OutFile,'Sheet',OutSheet,'Range',['A' num2str(existing_rows+1)],'WriteVariableNames',false);


%% Output removal data with depth if desired
if OutputRemoval==1
    switch IRIS_Type
        case 'Mn'
            RemovalHeaders={'Depth (in)','Average Mn Removal','Avg Removal with Fe Pixels as 100% Removal'};
            Depthvar=[mean_I_L_diff_y,mean_I_L_diff_y_noFe];
        case 'Fe'
            RemovalHeaders={'Depth (in)','Average Fe Removal'};
            Depthvar=mean_I_L_diff_y;
        case 'S'
            RemovalHeaders={'Depth (in)','Average S (relative intensity)'};
            Depthvar=mean_S_y;
    end
        
   %check to see if sheet exists
   output_sheets=sheetnames(OutFile);
   if sum(strcmp(output_sheets,End_file))
       existing_output_removal=readtable(OutFile,'Sheet',End_file);     
       existing_rows_removal=1+size(existing_output_removal,1);
   else
       existing_rows_removal=0;
   end

   if OutputRemovalDepth==0
       Removal_Avg=[y_length',Depthvar];
    elseif OutputRemovalDepth>0
        In_per_pixel=length(y_length)/y_length(length(y_length)); %ratio of inches per pixel
        Inc=floor(OutputRemovalDepth*In_per_pixel);%increment
        num_depths=floor(y_length(length(y_length))/OutputRemovalDepth); %number of depths to calculate
        Removal_Avg=zeros(num_depths,size(RemovalHeaders,2)); %Initialize output array
        for i=1:num_depths %perform averaging
            start_index=floor((i-1)*OutputRemovalDepth*In_per_pixel)+1;
            end_index=floor(i*OutputRemovalDepth*In_per_pixel);
            Removal_Avg(i,1)=mean(y_length(start_index:end_index));
            Removal_Avg(i,2:size(Depthvar,2)+1)=mean(Depthvar(start_index:end_index,:));
        end
   end
   %Append data to sheet
   OutputTableDepth=array2table(Removal_Avg,'VariableNames',RemovalHeaders);
   writetable(OutputTableDepth,OutFile,'Sheet',End_file,'Range',['A' num2str(existing_rows_removal+1)]);

end

        

beep
pause(1)
beep
Messages{size(Messages,1)+1,1}=(sprintf('Program run complete!\n'));
app.OutputText.Value=Messages;
 
catch e
for i=1:size(e.stack)
    e.stack(i,:).name
    e.stack(i).line
    Messages{size(Messages,1)+1,1}=sprintf(['Error in ', e.stack(i,:).name, ', line: ',num2str(e.stack(i).line)]);
end
Messages{size(Messages,1)+1,1}=sprintf(e.identifier);
Messages{size(Messages,1)+1,1}=sprintf(e.message);

app.OutputText.Value=Messages;

end

%% Functions
function [Circles] = CIRCLE_ID(fighan)
%CIRCLE_HUNTER User selects the center of the cirlce
%   fighan is the figure handle to zoom on
title('Drag a box around a circle, press SPACE');
figure(fighan)
zoom on
while ~waitforbuttonpress 
    end
title('Select center of circle and press SPACE');
dcm_obj = datacursormode; %Enable point selection
set(dcm_obj,'DisplayStyle','datatip',...
    'SnapToDataVertex','off','Enable','on')
while ~waitforbuttonpress 
    end
Point = getCursorInfo(dcm_obj); %ID point coordinates
Circles = Point.Position; %ID point coordinates
zoom out
hold on
plot(Circles(1),Circles(2),'r+');
zoom off
title('Circle selection completed');
end

function [Center,Radii] = CIRCLE_HUNTER(Image,HuntRadii,polarity)
%CIRCLE_HUNTER Iterative version of imfindcircles
%   Image is the image containing 1 circle
%   HuntRadii is a vector of the min and max circle radius to hunt for
%   polarity is either 'bright' or 'dark'
Sensitivity=.85; %Circle detection sensitivity
multiple_circle_counter=1;
for i=1:50
    fprintf(' Iteration %d: \t',i);
    [Center, Radii] = imfindcircles(Image, HuntRadii,'Sensitivity',Sensitivity,'ObjectPolarity',polarity);
    fprintf(' Number of circles detected: %d\n',length(Radii));
    if length(Radii)==1
        break;
    elseif length(Radii)<1  %adjust sensitivity to find cirlces
        Sensitivity=Sensitivity+0.007;
    else
        Sensitivity=Sensitivity-0.01;
        Center_x_vec(multiple_circle_counter:multiple_circle_counter+length(Radii)-1)=Center(:,1);
        Center_y_vec(multiple_circle_counter:multiple_circle_counter+length(Radii)-1)=Center(:,2);
        Radii_vec(multiple_circle_counter:multiple_circle_counter+length(Radii)-1)=Radii;
        multiple_circle_counter=multiple_circle_counter+length(Radii);
    end
    
    if Sensitivity >1 || Sensitivity <0
        fprintf('Sensitivity exceeds bounds for initial circle hunting\n');
        return;
    end
    if i==50
        if multiple_circle_counter==1
            fprintf('No circles found. Choosing the middle of the subimage!\n');
            Center=[size(Image,1)/2, size(Image,2)/2];
            Radii=mean(HuntRadii);
        else
            fprintf('Iteration limit reached for initial circle hunting. Taking an average of found circles.\n');
            Center=[mean(Center_x_vec),mean(Center_y_vec)];
            Radii=mean(Radii_vec);
        end
        return;
    end
end
end

function [Corners] = CORNER_ID(fighan)
%CORNER_HUNTER User selects the corner points
%   fighan is the figure handle to zoom on
Points=4;
Corners=ones(Points,2); %initialize matrix

for i=1:Points
    if i==1
        title('Drag a box around top left corner, press SPACE');
    elseif i==2
        title('Drag a box around top right corner, press SPACE');
    elseif i==3
        title('Drag a box around bottom right corner, press SPACE');
    else
        title('Drag a box around bottom left corner, press SPACE');
    end
    figure(fighan)
    zoom on
    while ~waitforbuttonpress 
    end
    title('Select corner and press SPACE');
    dcm_obj = datacursormode; %Enable point selection
    set(dcm_obj,'DisplayStyle','datatip',...
        'SnapToDataVertex','off','Enable','on')
    while ~waitforbuttonpress 
    end
    Point = getCursorInfo(dcm_obj); %ID point coordinates
    Corners(i,:) = Point.Position; %ID point coordinates
    zoom out
    hold on
    plot(Corners(i,1),Corners(i,2),'r+');
end
zoom off
title('Corner selection completed');
end

function CIELAB_PLOT(Point1,Point2)
%CIELAB_PLOT Plots CIELAB Color Space
% From cielabplot.m
% Stephen Westland (2020). Computational Colour Science using MATLAB 2e 
%(https://www.mathworks.com/matlabcentral/fileexchange/40640-computational-colour-science-using-matlab-2e),
% MATLAB Central File Exchange. Retrieved June 16, 2020.
%   Point1 - Point to plot in LAB [a b] from 0 to 1 (optional)
%   Point2 - Point to plot in LAB [a b] from 0 to 1 (optional)

scaling=1; %scaling factor to adjust units

plot([0 0],[-60 60],'k','LineWidth',2);
hold on;
plot([-60 60],[0 0],'k','LineWidth',2);
axis equal;
axis([-60 60 -60 60]);
gr=[.7 .7 .7];
r=[.9 0 0];
g=[0 .9 0];
y=[.9 .9 0];
bl=[0 0 .9];
index=0;

%first quadrant
index=index+1;
a=50;
b=0;
ab(index,:)=[a b];
col(index,:)=r;
for i=5:5:85
    index=index+1;
    a=cos(i*pi/180)*50;
    b=sin(i*pi/180)*50;
    ab(index,:)=[a b];
    c=(a*r+(50-a)*y)/50;
    col(index,:)=c;
end
index=index+1;
a=0;
b=50;
ab(index,:)=[a b];
col(index,:)=y;

%grey
index=index+1;

a=0;
b=0;
ab(index,:)=[a b];
col(index,:)=gr;

patch('Vertices',ab,'Faces',1:size(ab,1),'EdgeColor','none','FaceVertexCData',col,'FaceColor','interp');
clear ab;

index=0;

%Second Quadrant
index=index+1;
a=0;
b=50;
ab(index,:)=[a b];
col(index,:)=y;

for i=95:5:175
    index=index+1;
    a=cos(i*pi/180)*50;
    b=sin(i*pi/180)*50;
    ab(index,:)=[a b];
    c=(b*y+(50-b)*g)/50;
    col(index,:)=c;
end
index=index+1;
a=-50;
b=0;
ab(index,:)=[a b];
col(index,:)=g;

%grey
index=index+1;

a=0;
b=0;
ab(index,:)=[a b];
col(index,:)=gr;

patch('Vertices',ab,'Faces',1:size(ab,1),'EdgeColor','none','FaceVertexCData',col,'FaceColor','interp');
clear ab;
index=0;

%Third Quadrant
index=index+1;
a=-50;
b=0;
ab(index,:)=[a b];
col(index,:)=g;

for i=185:5:265
    index=index+1;
    a=cos(i*pi/180)*50;
    b=sin(i*pi/180)*50;
    ab(index,:)=[a b];
    c=(-b*bl+(50+b)*g)/50;
    col(index,:)=c;
end
index=index+1;
a=0;
b=-50;
ab(index,:)=[a b];
col(index,:)=bl;
%grey
index=index+1;

a=0;
b=0;
ab(index,:)=[a b];
col(index,:)=gr;

patch('Vertices',ab,'Faces',1:size(ab,1),'EdgeColor','none','FaceVertexCData',col,'FaceColor','interp');
clear ab;
index=0;

% Fourth Quadrant
index=index+1;
a=0;
b=-50;
ab(index,:)=[a b];
col(index,:)=bl;

for i=275:5:355
    index=index+1;
    a=cos(i*pi/180)*50;
    b=sin(i*pi/180)*50;
    ab(index,:)=[a b];
    c=(a*r+(50-a)*bl)/50;
    col(index,:)=c;
end
index=index+1;
a=50;
b=0;
ab(index,:)=[a b];
col(index,:)=r;

%grey
index=index+1;

a=0;
b=0;
ab(index,:)=[a b];
col(index,:)=gr;

patch('Vertices',ab,'Faces',1:size(ab,1),'EdgeColor','none','FaceVertexCData',col,'FaceColor','interp');
clear ab;

plot([0 0],[-60 60],'k','LineWidth',2);
plot([-60 60],[0 0],'k','LineWidth',2);

xlabel('a*');
ylabel('b*');

if nargin==1
    plot(Point1(1)*scaling,Point1(2)*scaling,'ok');
    hold on;
elseif nargin==2
    plot(Point1(1)*scaling,Point1(2)*scaling,'ok');
    hold on;
    plot(Point2(1)*scaling,Point2(2)*scaling,'dk');
    title('Initial Point: circle   Final Point: diamond')
end
end

function [Color_Match] = MUNSELL_MATCH(C_in)
%MUNSELL_MATCH Matches image color to Munsell soil colors
%   C_in is a matrix of [L*,A*,B*] to be matched in L*A*B* color space
%   where the number of rows is the number of colors to be matched

%% Constants
n_colors=size(C_in,1);

%% Munsell Colors (Scanned in)
ParamFile='AdjustableParameters.xlsx'; %Parameter file name
Munsell=readtable(ParamFile,'Sheet','Munsell');
Munsell_Names=Munsell.Color;
Munsell_LAB=[Munsell.L, Munsell.A, Munsell.B];

%% Initialize Matrices
n_Munsell=size(Munsell_LAB,1);
Color_Match=strings(n_colors,1);

%% Match colors
for i=1:n_colors
    All_Distances=zeros(n_Munsell,1);
    for j=1:n_Munsell
        All_Distances(j)=CIEDE2000(C_in(i,:),Munsell_LAB(j,:));
    end
    [~,Index]=min(All_Distances);
    Color_Match(i)=Munsell_Names{Index};
end
end

function distance = CIEDE2000 (Color1, Color2)
%CIEDE2000 calculates the CIE delta E2000 color difference
L1=Color1(1);
A1=Color1(2);
B1=Color1(3);
L2=Color2(1);
A2=Color2(2);
B2=Color2(3);
% L1, A1, and B1, are the L*A*B* values from the first color
% L2, A2, and B2, are the L*A*B* values from the second color
% The function returns the squared distance between the 2 points
% https://en.wikipedia.org/wiki/Color_difference

k_L=1;
k_C=1;
k_H=1;

delLprime=L2-L1;
Lbar=(L2+L1)/2;

C1=sqrt(A1^2+B1^2);
C2=sqrt(A2^2+B2^2);
Cbar=(C1+C2)/2;

A1prime=A1+A1/2*(1-sqrt(Cbar^7/(Cbar^7+25)));
A2prime=A2+A2/2*(1-sqrt(Cbar^7/(Cbar^7+25)));

C1prime=sqrt(A1prime^2+B1^2);
C2prime=sqrt(A2prime^2+B2^2);

Cbarprime=(C1prime+C2prime)/2;
delCprime=C2prime-C1prime;

if A1prime==0 && B1==0
	h1prime=0;
else
	h1prime=atan2d(B1,A1prime)+180; %shift from -180:180 to 0:360 degrees
end

if A2prime==0 && B2==0
	h2prime=0;
else
	h2prime=atan2d(B2,A2prime)+180;
end

if h1prime==0 || h2prime == 0
	delhprime=0;
	Hbarprime=h1prime+h2prime;
elseif abs(h1prime-h2prime)<=180
	delhprime=h2prime-h1prime;
	Hbarprime=(h1prime+h2prime)/2;
elseif abs(h1prime-h2prime)>180 && h2prime<=h1prime
	delhprime=h2prime-h1prime+360;
else
	delhprime=h2prime-h1prime-360;
end

if abs(h1prime-h2prime)>180 && h1prime+h2prime<360
	Hbarprime=(h1prime+h2prime+360)/2;
elseif abs(h1prime-h2prime)>180 && h1prime+h2prime>=360
	Hbarprime=(h1prime+h2prime-360)/2;
end

delHprime=2*sqrt(C1prime*C2prime)*sind(delhprime/2);

T=1-0.17*cosd(Hbarprime-30)+0.24*cosd(2*Hbarprime)+0.32*cosd(3*Hbarprime+6)-0.2*cosd(4*Hbarprime-63);

S_L=1+0.015*(Lbar-50)^2/sqrt(20+(Lbar-50)^2);
S_C=1+0.045*Cbarprime;
S_H=1+0.015*Cbarprime*T;

R_T=-2*sqrt(Cbarprime^7/(Cbarprime^7+25^7))*sind(60*exp(-1*((Hbarprime-275)/25)^2));

delE00_squared=(delLprime/(k_L*S_L))^2 + (delCprime/(k_C*S_C))^2 + (delHprime/(k_H*S_H))^2 + R_T*delCprime/(k_C*S_C)*delHprime/(k_H*S_H);

distance=delE00_squared;
end