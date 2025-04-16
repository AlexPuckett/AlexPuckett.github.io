%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%MIT License
%%
%%Copyright (c) 2025 Alexandros Puckett Anastasiou
%%
%%Permission is hereby granted, free of charge, to any person obtaining a copy
%%of this software and associated documentation files (the “Software”), to deal
%%in the Software without restriction, including without limitation the rights
%%to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
%%copies of the Software, and to permit persons to whom the Software is
%%furnished to do so, subject to the following conditions:
%%
%%The above copyright notice and this permission notice shall be included in all
%%copies or substantial portions of the Software.
%%
%%THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
%%IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
%%FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
%%AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
%%LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
%%OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
%%SOFTWARE.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



clearvars
clc

%% Creating Initial layout
% Creating the main UI figure
mainFig = uifigure("Name", "BrachyBarrier");
mainFig.Resize = 'off'; % Disable resizing
mainFig.WindowState = 'normal'; % Ensure the window is in normal state
set(mainFig, 'Units', 'normalized', 'Position', [0.05, 0.05, 0.9, 0.9]);

% Creating tab group and tab
tabGroup = uitabgroup(mainFig, 'Units', 'normalized', "Position", [0, 0, 1, 1]);
shieldingTab = uitab(tabGroup, "Title", "Shielding Thickness");
shieldingTab.Scrollable = "on";

%Panels
workloadPanel = uipanel(shieldingTab, 'Title', 'Workload', 'Position', [60, 400, 150, 200], 'BackgroundColor',[0.8 0.8 0.8], 'FontWeight', 'bold');
distancesPanel = uipanel(shieldingTab, 'Title', 'Wall Distances from Source', 'Position', [215, 400, 210, 200], 'BackgroundColor',[0.8 0.8 0.8], 'FontWeight', 'bold');
designparameteresPanel = uipanel(shieldingTab, 'Title', 'Design Parameters', 'Position', [655, 400, 300, 200], 'BackgroundColor',[0.8 0.8 0.8], 'FontWeight', 'bold', 'Scrollable', 'on');
areasPanel = uipanel(shieldingTab, 'Title', 'Wall Areas', 'Position', [430, 400, 220, 200], 'BackgroundColor',[0.8 0.8 0.8], 'FontWeight', 'bold');
anglesPanel = uipanel(shieldingTab, 'Title', 'Angles', 'Position', [960, 515, 180, 85], 'BackgroundColor',[0.8 0.8 0.8], 'FontWeight', 'bold');
pricesPanel = uipanel(shieldingTab, 'Title', 'Material Prices', 'Position', [960, 400, 180, 110], 'BackgroundColor',[0.8 0.8 0.8], 'FontWeight', 'bold');

%% Setting up source characteristics
% Source data
sourceData = struct();
sourceData.Ir192.RAKR = 0.111; %uGym^2/MBqh
sourceData.Ir192.E = 0.37; %MeV gamma ray
sourceData.Ir192.Lead = 28354; %mass att. coefficient mm2/kg through interpolation with values from NIST XCOM
sourceData.Ir192.Steel = 9877;
sourceData.Ir192.Concrete = 10139.1;
sourceData.Ir192.F = 0.93; % Tissue Air ratio F(D,theta) for path length of 10cm in water Safety Report Series No.47
sourceData.Ir192.TVLe = struct('Lead', 16, 'Steel', 43, 'Concrete', 152);
sourceData.Ir192.TVL1 = struct('Lead', [], 'Steel', 49, 'Concrete', []);

sourceData.Co60.RAKR = 0.308;
sourceData.Co60.E = 1.25;
sourceData.Co60.Lead = 5876;
sourceData.Co60.Steel = 5350;
sourceData.Co60.Concrete = 5807;
sourceData.Co60.F = 0.81;
sourceData.Co60.TVLe = struct('Lead', 41, 'Steel', 71, 'Concrete', 218);
sourceData.Co60.TVL1 = struct('Lead', [], 'Steel', 87, 'Concrete', 245);

sourceData.Cs137.RAKR = 0.077;
sourceData.Cs137.E = 0.662;
sourceData.Cs137.Lead = 11400;
sourceData.Cs137.Steel = 7390;
sourceData.Cs137.Concrete = 7923.21;
sourceData.Cs137.F = 0.86;
sourceData.Cs137.TVLe = struct('Lead', 22, 'Steel', 53, 'Concrete', 175);
sourceData.Cs137.TVL1 = struct('Lead', [], 'Steel', 69, 'Concrete', []);

sourceData.Au198.RAKR = 0.056;
sourceData.Au198.E = 0.42;
sourceData.Au198.Lead = 21812;
sourceData.Au198.Steel = 9203;
sourceData.Au198.Concrete = 9609.4;
sourceData.Au198.F = 0.90;
sourceData.Au198.TVLe = struct('Lead', 11, 'Steel', 0, 'Concrete', 142);
sourceData.Au198.TVL1 = struct('Lead', [], 'Steel', [], 'Concrete', []);

sourceData.Ra226.RAKR = 0.195;
sourceData.Ra226.E = 0.78;
sourceData.Ra226.Lead = 9231;
sourceData.Ra226.Steel = 7061.4;
sourceData.Ra226.Concrete = 7327.9;
sourceData.Ra226.F = 0.86;
sourceData.Ra226.TVLe = struct('Lead', 45, 'Steel', 76, 'Concrete', 240);
sourceData.Ra226.TVL1 = struct('Lead', [], 'Steel', 86, 'Concrete', []);

%% Setting up material characteristics
%Densities kg/mm3
density = struct();
density.Lead = 1.13e-5;
density.Steel = 7.9e-6;
density.Concrete = 2.35e-6;

Parameters = readtable('Materials.xlsx','Sheet','Parameters','ReadVariableNames',false);
massattcoef = struct();
massattcoef.Lead = readtable('Materials.xlsx','Sheet','Lead','ReadVariableNames',false);
massattcoef.Steel = readtable('Materials.xlsx','Sheet','Steel','ReadVariableNames',false);
massattcoef.Concrete = readtable('Materials.xlsx','Sheet','Concrete','ReadVariableNames',false);
massattcoef.Air = readtable('Materials.xlsx','Sheet','Air','ReadVariableNames',false);

%% Creating Labels, Edit Fields, Checkboxes and DropDowns
% Dropdown for sources
sourceLabel = uilabel(shieldingTab, 'Text', 'S', 'Interpreter', 'tex', 'Position', [20, 640, 10, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
sourceDropdown = uidropdown(shieldingTab, "Items", fieldnames(sourceData), 'Position', [35, 640, 65, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
sourceDropdown.ValueChangedFcn = @(dd, event) updateSourceData(dd, sourceData);

%Chechboxes
cbx = uicheckbox(shieldingTab,"Text","Pt.Att.", 'Position',[110,640,70,22]);
cbx_idr = uicheckbox(shieldingTab,"Text","IDR", 'Position',[175,640,70,22]);
cbx_nomaze = uicheckbox(shieldingTab,"Text","NoMaze", 'Position',[230,640,70,22]);
cbx_onelegmaze = uicheckbox(shieldingTab,"Text","OneLeg", 'Position',[300,640,70,22]);
cbx_twolegmaze = uicheckbox(shieldingTab,"Text","TwoLeg", 'Position',[370,640,70,22]);

% Labels and Edit Fields for workload calculations
activityLabel = uilabel(workloadPanel, 'Text', 'A[MBq]', 'Position', [10, 150, 75, 22], 'Interpreter', 'tex', 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
activityEditField = uieditfield(workloadPanel, 'numeric', 'Position', [85, 150, 48, 22], "ValueDisplayFormat", "%.2f");

numberSourcesLabel = uilabel(workloadPanel, 'Text', '#S', 'Interpreter', 'tex', 'Position', [10, 125, 75, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
numberSourcesEditField = uieditfield(workloadPanel, 'numeric', 'Position', [85, 125, 48, 22], "ValueDisplayFormat", "%.2f");

durationLabel = uilabel(workloadPanel, 'Text', 't [h]', 'Position', [10, 100, 75, 22], 'Interpreter', 'tex', 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
durationEditField = uieditfield(workloadPanel, 'numeric', 'Position', [85, 100, 48, 22], "ValueDisplayFormat", "%.2f");

treatmentsLabel = uilabel(workloadPanel, 'Text', 'Tr_{per week}', 'Interpreter', 'tex', 'Position', [10, 75, 75, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
treatmentsEditField = uieditfield(workloadPanel, 'numeric', 'Position', [85, 75, 48, 22], "ValueDisplayFormat", "%.2f", 'Editable', 'on');

% Labels and Edit Fields for Transmission Factor Calculations
distanceLabel = cell(1,6);
areaLabel = cell(1,6);
designLimitLabel = cell(1,29);
distanceEditField = cell(1,6);
areaEditField = cell(1,6);
for i = 1:6
    distanceLabel{i} = uilabel(distancesPanel, 'Text', ['d_{' num2str(i) '}[m]'], 'Interpreter', 'tex', 'Position', [10, 150-(i-1)*25, 75, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    areaLabel{i} = uilabel(areasPanel, 'Text', ['A_{' num2str(i) '}[m^{2}]'], 'Interpreter', 'tex', 'Position', [10, 150-(i-1)*25, 50, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    designLimitLabel{i} = uilabel(designparameteresPanel, 'Text', 'P[μGy]', 'Interpreter', 'tex', 'Position', [10, 710-(i-1)*25, 40, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    distanceEditField{i} = uieditfield(distancesPanel, 'numeric', 'Position', [50, 150-(i-1)*25, 48, 22], "ValueDisplayFormat", "%.2f");
    areaEditField{i} = uieditfield(areasPanel, 'numeric', 'Position', [50, 150-(i-1)*25, 48, 22], "ValueDisplayFormat", "%.2f");
end
for i = 7:9
    designLimitLabel{i} = uilabel(designparameteresPanel, 'Text', ['P_m{' num2str(i-6) '}'], 'Interpreter', 'tex', 'Position', [10, 560-(i-7)*25, 40, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
end
for i = 10:21
    designLimitLabel{i} = uilabel(designparameteresPanel, 'Text', ['P_e{' num2str(i-9) '}'], 'Interpreter', 'tex', 'Position', [10, 485-(i-10)*25, 40, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
end
for i = 22:29
    designLimitLabel{i} = uilabel(designparameteresPanel, 'Text', ['P_c{' num2str(i-21) '}'], 'Interpreter', 'tex', 'Position', [10, 185-(i-22)*25, 40, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
end

occupationFactorLabel = cell(1,29);
designLimitAreaDropdown = cell(1,29);
designLimitEditField = cell(1,29);
occupationFactorEditField = cell(1,29);
cbx_contamination = cell(1,29);
for i = 1:29
    occupationFactorLabel{i} = uilabel(designparameteresPanel, 'Text', 'T', 'Interpreter', 'tex', 'Position', [210, 710-(i-1)*25, 30, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    designLimitAreaDropdown{i} = uidropdown(designparameteresPanel, "Items", ["Select", "Controlled Area", "Uncontrolled Area", "Public Area"], 'Position', [50, 710-(i-1)*25, 65, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    designLimitEditField{i} = uieditfield(designparameteresPanel, 'numeric', 'Position', [120, 710-(i-1)*25, 48, 22], "ValueDisplayFormat", "%.2f", 'Editable', 'off');
    occupationFactorEditField{i} = uieditfield(designparameteresPanel, 'numeric', 'Position', [220, 710-(i-1)*25, 48, 22], "ValueDisplayFormat", "%.2f");
    cbx_contamination{i} = uicheckbox(designparameteresPanel,"Text","C",'Position',[170, 710-(i-1)*25, 30, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');

    % Callback to designLimitAreaDropdown to update designLimitEditField
    designLimitAreaDropdown{i}.ValueChangedFcn = @(dd, event) setDesignLimit(dd, designLimitEditField{i}, cbx_idr);

    %Callback to change the design limit if contamination factor is toggled
    cbx_contamination{i}.ValueChangedFcn = @(cbx, event) updateDesignLimitWithContamination(cbx, designLimitAreaDropdown{i}, designLimitEditField{i}, cbx_idr);
end

entrancedistLabel = cell(1,2);
entrancedistEditField = cell(1,2);
for i = 1:2
    entrancedistLabel{i} = uilabel(distancesPanel, 'Text', ['d_{d' num2str(i) '} [m]'], 'Interpreter', 'tex', 'Position', [110, 150-(i-1)*25, 60, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    entrancedistEditField{i} = uieditfield(distancesPanel, 'numeric', 'Position', [150, 150-(i-1)*25, 48, 22], "ValueDisplayFormat", "%.2f");
end
mazedistEditField = cell(1,3);
for i = 1:3
    mazedistLabel = uilabel(distancesPanel, 'Text', ['d_{m' num2str(i) '}[m]'], 'Interpreter', 'tex', 'Position', [110, 100-(i-1)*25, 60, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    mazedistEditField{i} = uieditfield(distancesPanel, 'numeric', 'Position', [150, 100-(i-1)*25, 48, 22], "ValueDisplayFormat", "%.2f");
end

mazeareaaEditField = cell(1,2);
mazeareabEditField = cell(1,2);
incidentangleEditField = cell(1,2);
refangleEditField = cell(1,2);
mazeareaLabel = uilabel(areasPanel, 'Text', 'A_{ma} [m^{2}]', 'Interpreter', 'tex', 'Position', [110, 150, 60, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
mazeareaEditField = uieditfield(areasPanel, 'numeric', 'Position', [160, 150, 48, 22], "ValueDisplayFormat", "%.2f");
mazeareaLabel_ = uilabel(areasPanel, 'Text', 'A_{mb} [m^{2}]', 'Interpreter', 'tex', 'Position', [110, 125, 60, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
mazeareaEditField_ = uieditfield(areasPanel, 'numeric', 'Position', [160, 125, 48, 22], "ValueDisplayFormat", "%.2f");
for i = 1:2
    mazeareaaLabel = uilabel(areasPanel, 'Text', ['A_{m' num2str(i) 'a}[m^{2}]'], 'Interpreter', 'tex', 'Position', [110, 100-(i-1)*50, 60, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    mazeareabLabel = uilabel(areasPanel, 'Text', ['A_{m' num2str(i) 'b}[m^{2}]'], 'Interpreter', 'tex', 'Position', [110, 75-(i-1)*50, 60, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    incidentangleLabel = uilabel(anglesPanel, 'Text', ['θ°_o' num2str(i)], 'Interpreter', 'tex', 'Position', [10+(i-1)*85, 35, 30, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    refangleLabel = uilabel(anglesPanel, 'Text', ['θ°_r' num2str(i)], 'Interpreter', 'tex', 'Position', [10+(i-1)*85, 10, 30, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
    mazeareaaEditField{i} = uieditfield(areasPanel, 'numeric', 'Position', [160, 100-(i-1)*50, 48, 22], "ValueDisplayFormat", "%.2f");
    mazeareabEditField{i} = uieditfield(areasPanel, 'numeric', 'Position', [160, 75-(i-1)*50, 48, 22], "ValueDisplayFormat", "%.2f");
    incidentangleEditField{i} = uieditfield(anglesPanel, 'numeric', 'Position', [35+(i-1)*85, 35, 48, 22], "ValueDisplayFormat", "%.2f");
    refangleEditField{i} = uieditfield(anglesPanel, 'numeric', 'Position', [35+(i-1)*85, 10, 48, 22], "ValueDisplayFormat", "%.2f");
end

workloadLabel = uilabel(workloadPanel, 'Text', 'W', 'Interpreter', 'tex', 'Position', [10, 20, 175, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
workloadValue = uilabel(workloadPanel, 'Text', '-', 'Position', [40, 20, 120, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');

% Labels and Edit Fields for Material Prices
LeadPriceLabel = uilabel(pricesPanel, "Text", "Pb Pr.[Eu/Kg]", 'Interpreter', 'tex', "Position", [10, 60, 100, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
SteelPriceLabel = uilabel(pricesPanel, "Text", "Stl Pr.[Eu/Kg]", 'Interpreter', 'tex', "Position", [10, 35, 100, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');
ConcretePriceLabel = uilabel(pricesPanel, "Text", "Conc Pr.[Eu/Kg]", 'Interpreter', 'tex', "Position", [10, 10, 110, 22], 'FontWeight', 'bold', 'FontSize', 10, 'FontColor', 'black');

PriceEditField = struct();
PriceEditField.Lead = uieditfield(pricesPanel, 'numeric', 'Position', [110, 60, 48, 22], "ValueDisplayFormat", "%.2f");
PriceEditField.Steel = uieditfield(pricesPanel, 'numeric', 'Position', [110, 35, 48, 22], "ValueDisplayFormat", "%.2f");
PriceEditField.Concrete = uieditfield(pricesPanel, 'numeric', 'Position', [110, 10, 48, 22], "ValueDisplayFormat", "%.2f");

%% Setting up tables to insert data
tableData = cell(3, 17);  % Cell array for the table data

% Column headers for the result table
columnNames = cell(1,17);
columnNames{1} = 'Materials';
columnNames{14} = 'MazeThickness1';
columnNames{15} = 'MazeCost1';
columnNames{16} = 'MazeThickness2';
columnNames{17} = 'MazeCost2';
for i = 1:6
    columnNames{2*i} = ['Thickness' num2str(i)];
    columnNames{2*i+1} = ['Cost' num2str(i)];
end
% Create the table in the UI (positioned at the bottom for displaying results)
resultTable = uitable(shieldingTab, 'Data', tableData, 'ColumnName', columnNames, 'Position', [10, 250, 1200, 117], 'ColumnWidth', repmat({76}, 1, 12));

shieldData = cell(4,29); % Cell array for the shielding data
%Column headers for shielding table
columnNames1 = cell(1,29);
columnNames1{1} = 'Shielding';
for i = 1:6
    columnNames1{i+1} = ['Distance' num2str(i)];
end
for i = 8:19
    columnNames1{i} = ['EdgeDistance' num2str(i-7)];
end
for i = 20:27
    columnNames1{i} = ['CornerDistance' num2str(i-19)];
end
for i = 28:29
    columnNames1{i} = ['Leg' num2str(i-27)];
end
shieldTable = uitable(shieldingTab, 'Data', shieldData, 'ColumnName', columnNames1, 'Position', [10, 130, 850, 117], 'ColumnWidth', repmat({95}, 1, 6));

%Column headers for maze entrance DR table
mazeData = cell(3,3);
columnNames2 = cell(1,3);
columnNames2{1} = 'Materials';
for i = 1:2
    columnNames2{i+1} = ['MazeLeg' num2str(i)];
end
mazeTable = uitable(shieldingTab, 'Data', mazeData, 'ColumnName', columnNames2, 'Position', [880, 140, 327, 99], 'ColumnWidth', repmat({95}, 1, 6));

%% Setting up Callback Functions
% Adding the Save button for exporting table data to an Excel file
saveButton = uibutton(shieldingTab, 'Position', [440, 640, 50, 22], 'Text', '💾');
saveButton.ButtonPushedFcn = @(btn, event) saveToExcel(resultTable, shieldTable, mazeTable, mainFig);

% Calculating shielding thickness and updating the table
calcButton = uibutton(shieldingTab, 'Position', [500, 640, 50, 22], 'Text', '►');
calcButton.ButtonPushedFcn = @(btn, event) calculateShieldingThickness(sourceDropdown, activityEditField, durationEditField, treatmentsEditField, workloadValue, designLimitEditField, distanceEditField, occupationFactorEditField, areaEditField, entrancedistEditField, mazedistEditField, mazeareaEditField, mazeareaEditField_, mazeareaaEditField, mazeareabEditField, incidentangleEditField, refangleEditField, numberSourcesEditField, sourceData, density, PriceEditField, resultTable, tableData, shieldTable, shieldData, mazeTable, mazeData, mainFig, cbx, cbx_idr, massattcoef, cbx_onelegmaze, cbx_twolegmaze, Parameters);

% Callback for cbx_idr checkbox to update design limits and toggle treatment fields
cbx_idr.ValueChangedFcn = @(src, event) IDRToggle(src, designLimitAreaDropdown, designLimitEditField, treatmentsEditField, occupationFactorEditField);

% Store checkboxes in an array
checkboxes = [cbx_nomaze, cbx_onelegmaze, cbx_twolegmaze];

% Assign the common callback to each checkbox in a loop
for i = 1:numel(checkboxes)
    checkboxes(i).ValueChangedFcn = @(src, event) MazeToggle(src, cbx_nomaze, cbx_onelegmaze, cbx_twolegmaze, mazedistEditField, mazeareaEditField, mazeareaEditField_, mazeareaaEditField, mazeareabEditField, incidentangleEditField, refangleEditField, entrancedistEditField);
end

%% Setting up toggles
% Function to toggle the editable state of treatment fields and the design Limit values
function IDRToggle(cbx_idr, designLimitAreaDropdown, designLimitEditField, treatmentsEditField, occupationFactorEditField)
% Update all design limit fields
for i = 1:length(designLimitAreaDropdown)
    % Call setDesignLimit for each dropdown and edit field pair
    setDesignLimit(designLimitAreaDropdown{i}, designLimitEditField{i}, cbx_idr);
    if cbx_idr.Value % Checkbox is selected
        treatmentsEditField.Editable = 'off';
        occupationFactorEditField{i}.Editable = 'off';
    else % Checkbox is not selected
        treatmentsEditField.Editable = 'on';
        occupationFactorEditField{i}.Editable = 'on';
    end
end
end

function MazeToggle(src, cbx_nomaze, cbx_onelegmaze, cbx_twolegmaze, mazedistEditField, mazeareaEditField, mazeareaEditField_, mazeareaaEditField, mazeareabEditField, incidentangleEditField, refangleEditField, entrancedistEditField)
if src == cbx_nomaze
    cbx_onelegmaze.Value = false;
    cbx_twolegmaze.Value = false;
elseif src == cbx_onelegmaze
    cbx_nomaze.Value = false;
    cbx_twolegmaze.Value = false;
elseif src == cbx_twolegmaze
    cbx_nomaze.Value = false;
    cbx_onelegmaze.Value = false;
end

for i = 1:3
    if cbx_nomaze.Value
        mazedistEditField{i}.Editable = 'off';
    elseif cbx_onelegmaze.Value
        mazedistEditField{1}.Editable = 'on';
        mazedistEditField{2}.Editable = 'on';
        mazedistEditField{3}.Editable = 'off';
    elseif cbx_twolegmaze.Value
        mazedistEditField{i}.Editable = 'on';
    end
end

for i = 1:2
    if cbx_nomaze.Value
        mazeareaEditField.Editable = 'off';
        mazeareaEditField_.Editable = 'off';
        mazeareaaEditField{i}.Editable = 'off';
        mazeareabEditField{i}.Editable = 'off';
        incidentangleEditField{i}.Editable = 'off';
        refangleEditField{i}.Editable = 'off';
        entrancedistEditField{i}.Editable = 'off';
    elseif cbx_onelegmaze.Value
        mazeareaEditField.Editable = 'on';
        mazeareaEditField_.Editable = 'off';
        mazeareaaEditField{1}.Editable = 'on';
        mazeareabEditField{1}.Editable = 'on';
        mazeareaaEditField{2}.Editable = 'off';
        mazeareabEditField{2}.Editable = 'off';
        incidentangleEditField{1}.Editable = 'on';
        incidentangleEditField{2}.Editable = 'off';
        refangleEditField{1}.Editable = 'on';
        refangleEditField{2}.Editable = 'off';
        entrancedistEditField{1}.Editable = 'on';
        entrancedistEditField{2}.Editable = 'off';
    elseif cbx_twolegmaze.Value
        mazeareaEditField.Editable = 'on';
        mazeareaEditField_.Editable = 'on';
        mazeareaaEditField{i}.Editable = 'on';
        mazeareabEditField{i}.Editable = 'on';
        incidentangleEditField{i}.Editable = 'on';
        refangleEditField{i}.Editable = 'on';
        entrancedistEditField{i}.Editable = 'on';
    end
end
end

% Function to update designLimitEditField based on dropdown selection
function setDesignLimit(dd, designLimitEditField, cbx_idr)
if cbx_idr.Value
    switch dd.Value
        case "Controlled Area"
            designLimitEditField.Value = 25 * (3/10); %uGy/h
        case "Uncontrolled Area"
            designLimitEditField.Value = 2.5 * (3/10);
        case "Public Area"
            designLimitEditField.Value = 0.5 * (3/10);
        otherwise
            designLimitEditField.Value = 0; % Default or clear value
    end
else
    % If IDR checkbox is not checked
    switch dd.Value
        case "Controlled Area"
            designLimitEditField.Value = (400 * (3/10))/40; % uGy/week
        case "Uncontrolled Area"
            designLimitEditField.Value = (120 * (3/10))/40;
        case "Public Area"
            designLimitEditField.Value = (20 * (3/10))/40;
        otherwise
            designLimitEditField.Value = 0; % Default or clear value
    end
end
end

% Function to update the design limit based on the contamination checkbox
function updateDesignLimitWithContamination(cbx_contamination, dd, designLimitEditField, cbx_idr)
% Call setDesignLimit to update based on area selection first
setDesignLimit(dd, designLimitEditField, cbx_idr);

% If contamination checkbox is checked, adjust the design limit
if cbx_contamination.Value
    designLimitEditField.Value = designLimitEditField.Value * (1/2);
end
end

%% Functions for calculations
% Update source data based on selected source
function selectedSource = updateSourceData(sourceDropdown, sourceData)
selectedSource = sourceData.(sourceDropdown.Value);
end

% Function to calculate shielding thickness and update the table
function calculateShieldingThickness(sourceDropdown, activityEditField, durationEditField, treatmentsEditField, workloadValue, designLimitEditField, distanceEditField, occupationFactorEditField, areaEditField, entrancedistEditField, mazedistEditField, mazeareaEditField, mazeareaEditField_, mazeareaaEditField, mazeareabEditField, incidentangleEditField, refangleEditField, numberSourcesEditField, sourceData, density, PriceEditField, resultTable, tableData, shieldTable, shieldData, mazeTable, mazeData, mainFig, cbx, cbx_idr, massattcoef, cbx_onelegmaze, cbx_twolegmaze, Parameters)

% Get the selected source data
selectedSource = sourceData.(sourceDropdown.Value);

%% CALCULATING DOSES at direct distances etc
% Pre-allocate arrays to store thickness and cost results
transmissionFactor = zeros(1,6);
attenuationFactor = zeros(1,6);
thickness = zeros(3,6);  % 3 materials (Lead, Steel, Concrete), 6 distances
cost = zeros(3,6);       % Same size for cost
DoseRate = zeros(3,6);
InDoseRate = zeros(1,6);
mazethickness = zeros(3,1);
mazethickness_ = zeros(3,1);
mazecost = zeros(3,1);
mazecost_ = zeros(3,1);

if cbx.Value == 0
    F = 1;
elseif cbx.Value == 1
    F = selectedSource.F;
end

% Check for negative values in input fields
if any([activityEditField.Value, numberSourcesEditField.Value, durationEditField.Value, treatmentsEditField.Value] < 0)
    uialert(mainFig, 'Input values cannot be negative.', 'Input Error');
    return; % Exit the function if negative values are found
end

if cbx_idr.Value
    workload = selectedSource.RAKR * activityEditField.Value * numberSourcesEditField.Value; %Do = RAKR*A*#S
    % Set workload to 0 if it is negative
    if workload < 0
        workload = 0;
    end
    workloadValue.Text = sprintf('%.2e μGym^{2}/hour', workload);
    workloadValue.Interpreter = 'tex';
else
    % W = RAKR*A*#S*t*N
    workload = selectedSource.RAKR * activityEditField.Value * numberSourcesEditField.Value * durationEditField.Value * (treatmentsEditField.Value/40);

    % Set workload to 0 if it is negative
    if workload < 0
        workload = 0;
    end
    workloadValue.Text = sprintf('%.2e μGym^{2}/week', workload*40);
    workloadValue.Interpreter = 'tex';
end

% Loop through the 6 distances and calculate thickness and cost
for i = 1:6
    if cbx_idr.Value
        transmissionFactor(i) = (designLimitEditField{i}.Value * F * distanceEditField{i}.Value^2) / (workload);
        if cbx_onelegmaze.Value
            entrancetransmissionFactor  = (designLimitEditField{7}.Value * F * entrancedistEditField{1}.Value^2) / (workload);
            entrancetransmissionFactor_  = 0;
        elseif cbx_twolegmaze.Value
            entrancetransmissionFactor  = (designLimitEditField{7}.Value * F * entrancedistEditField{1}.Value^2) / (workload);
            entrancetransmissionFactor_  = (designLimitEditField{8}.Value * F * entrancedistEditField{2}.Value^2) / (workload);
        else
            entrancetransmissionFactor  = 0;
            entrancetransmissionFactor_  = 0;
        end
    else
        transmissionFactor(i) = (designLimitEditField{i}.Value * F * distanceEditField{i}.Value^2) / ((workload) * occupationFactorEditField{i}.Value);
        if cbx_onelegmaze.Value
            entrancetransmissionFactor  = (designLimitEditField{7}.Value * F * entrancedistEditField{1}.Value^2) / ((workload) * occupationFactorEditField{7}.Value);
            entrancetransmissionFactor_  = 0;
        elseif cbx_twolegmaze.Value
            entrancetransmissionFactor  = (designLimitEditField{7}.Value * F * entrancedistEditField{1}.Value^2) / ((workload) * occupationFactorEditField{7}.Value);
            entrancetransmissionFactor_  = (designLimitEditField{8}.Value * F * entrancedistEditField{2}.Value^2) / ((workload) * occupationFactorEditField{8}.Value);
        else
            entrancetransmissionFactor  = 0;
            entrancetransmissionFactor_  = 0;
        end
    end
    % Check if transmissionFactor is valid; if not, set to 0
    if isnan(transmissionFactor(i)) || transmissionFactor(i) == Inf || transmissionFactor(i) == -Inf || transmissionFactor(i) < 0
        transmissionFactor(i) = 0;
    end
    if isnan(entrancetransmissionFactor) || entrancetransmissionFactor == Inf || entrancetransmissionFactor == -Inf || entrancetransmissionFactor < 0
        entrancetransmissionFactor = 0;
    end
    if isnan(entrancetransmissionFactor_) || entrancetransmissionFactor_ == Inf || entrancetransmissionFactor_ == -Inf || entrancetransmissionFactor_ < 0
        entrancetransmissionFactor_ = 0;
    end

    attenuationFactor(i) = log10(1 / transmissionFactor(i));
    if cbx_onelegmaze.Value
        entranceattenutationFactor = log10(1/entrancetransmissionFactor);
        entranceattenutationFactor_ = 0;
    elseif cbx_twolegmaze.Value
        entranceattenutationFactor = log10(1/entrancetransmissionFactor);
        entranceattenutationFactor_ = log10(1/entrancetransmissionFactor_);
    else
        entranceattenutationFactor = 0;
        entranceattenutationFactor_ = 0;
    end

    % For each material, calculate thickness and cost
    materials = fieldnames(selectedSource.TVLe);  % Lead, Steel, Concrete
    for j = 1:length(materials)
        material = materials{j};
        TVLe = selectedSource.TVLe.(material);
        TVL1 = selectedSource.TVL1.(material);

        if isempty(TVL1)
            TVL1 = 0;  % Set TVL1 to zero if not provided
        end

        % Calculate thickness: thickness = TVL1 + (attenuationFactor - 1) * TVLe
        thickness(j,i) = TVL1 + (attenuationFactor(i)-1)*TVLe;
        if cbx_onelegmaze.Value
            mazethickness(j,1) = TVL1 + (entranceattenutationFactor - 1)* TVLe;
            mazethickness_(j,1) = 0;
        elseif cbx_twolegmaze.Value
            mazethickness(j,1) = TVL1 + (entranceattenutationFactor - 1)* TVLe;
            mazethickness_(j,1) = TVL1 + (entranceattenutationFactor_ - 1)* TVLe;
        else
            mazethickness(j,1) = 0;
            mazethickness_(j,1) = 0;
        end

        % Check if thickness is valid; if not, set to 0
        if isnan(thickness(j,i)) || thickness(j,i) == Inf || thickness(j,i) == -Inf
            thickness(j,i) = 0;
        end
        if isnan(mazethickness(j,1)) || mazethickness(j,1) == Inf || mazethickness(j,1) == -Inf
            mazethickness(j,1) = 0;
        end
        if isnan(mazethickness_(j,1)) || mazethickness_(j,1) == Inf || mazethickness_(j,1) == -Inf
            mazethickness_(j,1) = 0;
        end

        % Calculate cost based on material density and price per kg
        % Assuming thickness is in mm
        thickness_mm = ceil(thickness(j,i));
        if cbx_onelegmaze.Value
            mazethickness_mm = ceil(mazethickness(j,1));
            mazethickness_mm_ = 0;
        elseif cbx_twolegmaze.Value
            mazethickness_mm = ceil(mazethickness(j,1));
            mazethickness_mm_ = ceil(mazethickness_(j,1));
        else
            mazethickness_mm = 0;
            mazethickness_mm_ = 0;
        end
        volume = areaEditField{i}.Value*1e6*thickness_mm;  % Volume in mm
        mazevolume = mazeareaEditField.Value*1e6*mazethickness_mm;
        mazevolume_ = mazeareaEditField_.Value*1e6*mazethickness_mm_;
        cost(j,i) = PriceEditField.(material).Value * density.(material) * volume;
        mazecost(j,1) = PriceEditField.(material).Value * density.(material) * mazevolume;
        mazecost_(j,1) = PriceEditField.(material).Value * density.(material) * mazevolume_;

        % Check if cost is valid; if not, set to 0
        if isnan(cost(j,i)) || cost(j,i) == Inf || cost(j,i) == -Inf
            cost(j,i) = 0;
        end
        if isnan(mazecost(j,1)) || mazecost(j,1) == Inf || mazecost(j,1) == -Inf
            mazecost(j,1) = 0;
        end
        if isnan(mazecost_(j,1)) || mazecost_(j,1) == Inf || mazecost_(j,1) == -Inf
            mazecost_(j,1) = 0;
        end


        %% Calculating Dose Rates for Direct Distances
        % Do = W/d^2
        InDoseRate(i) = (workload)/(distanceEditField{i}.Value^2);

        % Check if Activity is valid; if not, set to 0
        if isnan(InDoseRate(i)) || InDoseRate(i) == Inf || InDoseRate(i) == -Inf
            InDoseRate(i) = 0;
        end

        DoseRate(j,i) = InDoseRate(i)*exp(-selectedSource.(material)*density.(material)*ceil(thickness(j,i)));

        % Check if Activity is valid; if not, set to 0
        if isnan(DoseRate(j, i)) || DoseRate(j, i) == Inf || DoseRate(j, i) == -Inf
            DoseRate(j, i) = 0;
        end

        % Initialize arrays with an additional dimension for each i
        mixDoseRate = zeros(1, 216, 6);
        mixCost = zeros(1, 216, 6); 
        proportions = zeros(216, 3, 6);

        % Initialize index for cells
        index = 1;
        for p1 = 0:0.2:1
            for p2 = 0:0.2:1
                for p3 = 0:0.2:1
                    % Calculate the mixed dose rate for each combination of proportions
                    mixDoseRate(index, 1, i) = InDoseRate(i) * exp(-((selectedSource.(materials{1}) * density.(materials{1}) * ceil(thickness(1,i)) * p1) + (selectedSource.(materials{2}) * density.(materials{2}) * ceil(thickness(2,i)) * p2) + (selectedSource.(materials{3}) * density.(materials{3}) * ceil(thickness(3,i)) * p3)));
                    % Calculate the mixed volume
                    volume1 = areaEditField{i}.Value * 1e6 * thickness(1,i) * p1;
                    volume2 = areaEditField{i}.Value * 1e6 * thickness(2,i) * p2;
                    volume3 = areaEditField{i}.Value * 1e6 * thickness(3,i) * p3;
                    % Calculate the mixed cost
                    mixCost(index, 1, i) = (PriceEditField.(materials{1}).Value * density.(materials{1})*volume1 + PriceEditField.(materials{2}).Value * density.(materials{2})*volume2 + PriceEditField.(materials{3}).Value * density.(materials{3})*volume3);
                    % Store proportions for this combination
                    proportions(index, :, i) = [p1, p2, p3];

                    % Increment the index
                    index = index + 1;
                    if index > 216
                        break;  % Exit loop if index exceeds the limit
                    end
                end
            end
        end

        % Reset bestDoseRate and bestCost for each iteration of i
        bestDoseRate = inf;
        bestCost = inf;
        bestProportions = [0, 0, 0];  % Initialize best proportions for each i

        % Find the best mixture based on dose rate and cost for this i
        for idx = 2:216
            if (mixDoseRate(idx, 1, i) <= 0.075) && (mixCost(idx, 1, i) <= bestCost)
                bestDoseRate = mixDoseRate(idx, 1, i);
                bestCost = mixCost(idx, 1, i);
                bestProportions = proportions(idx, :, i);  % Store the best proportions for this i
            end
        end

        % Store the best results for each iteration i
        tableData{4,1} = sprintf('Best Mix');
        tableData{4,2*i} = sprintf('%.2f, %.2f, %.2f', bestProportions(1), bestProportions(2), bestProportions(3));
        tableData{4,2*i+1} = sprintf('€ %.2e', bestCost);
        shieldData{5,1} = sprintf('Best Mix');
        shieldData{5,i+1} = sprintf('%.2e uSv/h', bestDoseRate);

        % Update the table data for the current material and distance
        tableData{j,1} = material;
        tableData{j,2*i} = sprintf('%.2f mm', ceil(thickness(j,i)));  % Thickness in mm
        tableData{j,2*i+1} = sprintf('€ %.2f', cost(j,i));        % Cost in EUR
        tableData{j,14} = sprintf('%.2f mm', ceil(mazethickness(j,1)));
        tableData{j,15} = sprintf('€ %.2e', mazecost(j,1));
        tableData{j,16} = sprintf('%.2f mm', ceil(mazethickness_(j,1)));
        tableData{j,17} = sprintf('€ %.2e', mazecost_(j,1));

        %Update the shield data for material and distance
        shieldData{1,1} = sprintf('No Shielding');
        shieldData{2,1} = sprintf('Lead Shield');
        shieldData{3,1} = sprintf('Steel Shield');
        shieldData{4,1} = sprintf('Concrete Shield');
        shieldData{1,i+1} = sprintf('%.2e uSv/h', InDoseRate(i));
        shieldData{j+1,i+1} = sprintf('%.2e uSv/h', DoseRate(j,i));

        s1 = uistyle('BackgroundColor','r');
        s2 = uistyle('BackgroundColor','y');
        s3 = uistyle('BackgroundColor','g');

        if cbx_idr.Value
            if InDoseRate(i) < 7.5 %uSv/h
                addStyle(shieldTable,s3,'cell',[1,i+1]);
            elseif InDoseRate(i) == 7.5
                addStyle(shieldTable,s2,'cell',[1,i+1]);
            else
                addStyle(shieldTable,s1,'cell',[1,i+1]);
            end
        else
            if InDoseRate(i) < 3 %uSv/h (weekly)
                addStyle(shieldTable,s3,'cell',[1,i+1]);
            elseif InDoseRate(i) == 3
                addStyle(shieldTable,s2,'cell',[1,i+1]);
            else
                addStyle(shieldTable,s1,'cell',[1,i+1]);
            end
        end

        if DoseRate(j,i) < designLimitEditField{i}.Value
            addStyle(shieldTable,s3,'cell',[j+1,i+1]);
            addStyle(shieldTable,s3,'cell',[5,i+1]);
        elseif DoseRate(j,i) == designLimitEditField{i}.Value
            addStyle(shieldTable,s2,'cell',[j+1,i+1]);
            addStyle(shieldTable,s2,'cell',[5,i+1]);
        else
            addStyle(shieldTable,s1,'cell',[j+1,i+1]);
            addStyle(shieldTable,s1,'cell',[5,i+1]);
        end
    end
end

%% CALCULATING DOSE RATES FOR EDGES
% Preallocate matrices
hypdist = zeros(4, 6); % Pairwise distances
hypth = zeros(length(materials), 4, 6);   % Pairwise thicknesses
hypInDoseRate = zeros(1, 24); % Dose rate at 12 points
hypDoseRate = zeros(length(materials), 24); % Dose rate per material at 12 points
besthypth = zeros(1,4,6);
besthypDoseRate = zeros(1,24);

% Calculate pairwise distances and thicknesses
idx = 1;
for i = 1:4
    for j = 1:6
        % Exclude invalid pairs
        if i == j || (i == 1 && j == 3) || (i == 2 && j == 1) || (i == 2 && j == 3) || (i == 2 && j == 4) || (i == 3 && j == 1) || (i == 4 && j == 1) || (i == 4 && j == 2) || (i == 4 && j == 3)
            continue;
        else
            for k = 1:length(materials)
                propString = tableData{4,2*j};
                propArray = sscanf(propString, '%f, %f, %f');
                prop1 = propArray(1);
                prop2 = propArray(2);
                prop3 = propArray(3);
                % Calculate distance and thickness
                hypdist(i, j) = sqrt(distanceEditField{i}.Value^2 + distanceEditField{j}.Value^2);
                hypth(k, i, j) = sqrt(thickness(k,i)^2 + thickness(k,j)^2);
                besthypth(1,i,j) = sqrt((thickness(1,i)*prop1+thickness(2,i)*prop2+thickness(3,i)*prop3)^2 + (thickness(1,j)*prop1+thickness(2,j)*prop2+thickness(3,j)*prop3)^2);

                % Calculate dose rate based on mode
                hypInDoseRate(idx) = (workload) / (hypdist(i, j)^2);

                if isnan(hypInDoseRate(idx)) || hypInDoseRate(idx) == Inf || hypInDoseRate(idx) == -Inf
                    hypInDoseRate(idx) = 0;
                end

                hypDoseRate(k, idx) = hypInDoseRate(idx) * exp(-selectedSource.(materials{k}) * density.(materials{k}) * ceil(hypth(k, i, j)));
                besthypDoseRate(1,idx) = hypInDoseRate(idx) * exp(-(selectedSource.(materials{1}) * density.(materials{1}) * ceil(besthypth(1, i, j))+selectedSource.(materials{2}) * density.(materials{2}) * ceil(besthypth(1, i, j))+selectedSource.(materials{1}) * density.(materials{1}) * ceil(besthypth(1, i, j))));
                if isnan(hypDoseRate(k,idx)) || hypDoseRate(k, idx) == Inf || hypDoseRate(k, idx) == -Inf
                    hypDoseRate(k,idx) = 0;
                end
            end
            idx = idx + 1;
            if idx > numel(hypInDoseRate)
                break;
            end
        end
    end
end

idx = 1;
for i = 8:19
    for j = 1:length(materials)
        shieldData{1,i} = sprintf('%.2e uSv/h', hypInDoseRate(idx));
        shieldData{j+1,i} = sprintf('%.2e uSv/h', hypDoseRate(j, idx));
        shieldData{5,i} = sprintf('%.2e uSv/h', besthypDoseRate(1, idx));
    end
    idx = idx + 1;
    if idx > numel(hypInDoseRate)
        break;
    end
end

s1 = uistyle('BackgroundColor','r');
s2 = uistyle('BackgroundColor','y');
s3 = uistyle('BackgroundColor','g');
for i = 8:19
    for idx = 1:12
        for j = 1:length(materials)
            if cbx_idr.Value
                if hypInDoseRate(idx) < 7.5 %uSv/h
                    addStyle(shieldTable,s3,'cell',[1,idx+7]);
                elseif hypInDoseRate(idx) == 7.5
                    addStyle(shieldTable,s2,'cell',[1,idx+7]);
                else
                    addStyle(shieldTable,s1,'cell',[1,idx+7]);
                end
            else
                if hypInDoseRate(idx) < 3 %uSv/h (weekly)
                    addStyle(shieldTable,s3,'cell',[1,idx+7]);
                elseif hypInDoseRate(idx) == 3
                    addStyle(shieldTable,s2,'cell',[1,idx+7]);
                else
                    addStyle(shieldTable,s1,'cell',[1,idx+7]);
                end
            end

            if hypDoseRate(j, idx) < designLimitEditField{idx+9}.Value
                addStyle(shieldTable,s3,'cell',[j+1,idx+7]);
                addStyle(shieldTable,s3,'cell',[5,idx+7]);
            elseif hypDoseRate(j, idx) == designLimitEditField{idx+9}.Value
                addStyle(shieldTable,s2,'cell',[j+1,idx+7]);
                addStyle(shieldTable,s2,'cell',[5,idx+7]);
            else
                addStyle(shieldTable,s1,'cell',[j+1,idx+7]);
                addStyle(shieldTable,s1,'cell',[5,idx+7]);
            end
        end
    end
end

%% CALCULATE DOSE RATES FOR CORNERS
chypdist = zeros(3,3,2);
chypth = zeros(length(materials),3,3,2);
chypInDoseRate = zeros(1,18);
chypDoseRate = zeros(length(materials),18);
bestchypth = zeros(1,3,3,2);
bestchypDoseRate = zeros(1,18);
idx = 1;
for i = 1:3
    for j = 2:4
        for l = 5:6
            if i == j || (i == 1 && j == 3) || (i == 2 && j == 4) || (i == 3 && j == 2)
                continue;
            else
                for k = 1:length(materials)
                    propString = tableData{4,2*l};
                    propArray = sscanf(propString, '%f, %f, %f');
                    prop1 = propArray(1);
                    prop2 = propArray(2);
                    prop3 = propArray(3);
                    chypdist(i,j,l) = sqrt(hypdist(i,j)^2 + distanceEditField{l}.Value^2);
                    chypth(k,i,j,l) = sqrt(hypth(k,i,j)^2 + thickness(k,l)^2);
                    bestchypth(1,i,j,l) = sqrt((besthypth(1,i,j))^2 + (thickness(1,l)*prop1+thickness(2,l)*prop2+thickness(3,l)*prop3)^2);
                    chypInDoseRate(idx) = (workload) / (chypdist(i,j,l)^2);

                    if isnan(chypInDoseRate(idx)) || chypInDoseRate(idx) == Inf || chypInDoseRate(idx) == -Inf
                        chypInDoseRate(idx) = 0;
                    end

                    chypDoseRate(k,idx) = chypInDoseRate(idx) * exp(-(selectedSource.(materials{1}) * density.(materials{1}) * ceil(chypth(1,i,j,l))+selectedSource.(materials{2}) * density.(materials{2}) * ceil(chypth(2,i,j,l))+selectedSource.(materials{3}) * density.(materials{3}) * ceil(chypth(3,i,j,l))));
                    bestchypDoseRate(1,idx) = hypInDoseRate(idx) * exp(-(selectedSource.(materials{1}) * density.(materials{1}) * ceil(bestchypth(1, i, j))+selectedSource.(materials{2}) * density.(materials{2}) * ceil(bestchypth(1, i, j))+selectedSource.(materials{1}) * density.(materials{1}) * ceil(bestchypth(1, i, j))));
                    if isnan(chypDoseRate(idx)) || chypDoseRate(idx) == Inf || chypDoseRate(idx) == -Inf
                        chypDoseRate(idx) = 0;
                    end
                end
                idx = idx + 1;
                if idx > numel(hypInDoseRate)
                    return;
                end
            end
        end
    end
end


idx = 1;
for i = 20:27
    for j = 1:length(materials)
        shieldData{1,i} = sprintf('%.2e uSv/h', chypInDoseRate(idx));
        shieldData{j+1,i} = sprintf('%.2e uSv/h', chypDoseRate(j, idx));
        shieldData{5,i} = sprintf('%.2e uSv/h', bestchypDoseRate(1, idx));
    end
    idx = idx + 1;
    if idx > numel(chypInDoseRate)
        break;
    end
end

s1 = uistyle('BackgroundColor','r');
s2 = uistyle('BackgroundColor','y');
s3 = uistyle('BackgroundColor','g');
for i = 20:27
    for idx = 1:8
        for j = 1:length(materials)
            if cbx_idr.Value
                if chypInDoseRate(idx) < 7.5 %uSv/h
                    addStyle(shieldTable,s3,'cell',[1,idx+19]);
                elseif chypInDoseRate(idx) == 7.5
                    addStyle(shieldTable,s2,'cell',[1,idx+19]);
                else
                    addStyle(shieldTable,s1,'cell',[1,idx+19]);
                end
            else
                if chypInDoseRate(idx) < 3 %uSv/h (weekly)
                    addStyle(shieldTable,s3,'cell',[1,idx+19]);
                elseif chypInDoseRate(idx) == 3
                    addStyle(shieldTable,s2,'cell',[1,idx+19]);
                else
                    addStyle(shieldTable,s1,'cell',[1,idx+19]);
                end
            end

            if chypDoseRate(j, idx) < designLimitEditField{idx+21}.Value
                addStyle(shieldTable,s3,'cell',[j+1,idx+19]);
                addStyle(shieldTable,s3,'cell',[5,idx+19]);
            elseif chypDoseRate(j, idx) == designLimitEditField{idx+21}.Value
                addStyle(shieldTable,s2,'cell',[j+1,idx+19]);
                addStyle(shieldTable,s2,'cell',[5,idx+19]);
            else
                addStyle(shieldTable,s1,'cell',[j+1,idx+19]);
                addStyle(shieldTable,s1,'cell',[5,idx+19]);
            end
        end
    end
end

%% CALCULATE DOSE RATES AT MAZE
electronrad = 2.82e-15;
E = selectedSource.E;
materials = fieldnames(selectedSource.TVLe);  % Lead, Steel, Concrete
MazeDoseRate = zeros(1,3);
MazeDoseRate1 = zeros(1,3);
scatangle = cell(1,2);
for j = 1:length(materials)
    material = materials{j};
    if cbx_onelegmaze.Value
        scatangle = 180 - (incidentangleEditField{1}.Value + refangleEditField{1}.Value);
        E_ = E/(1+((E/0.511)*(1-cosd(scatangle))));
        u1 = interp1(massattcoef.Air{:,1},massattcoef.Air{:,2},E);
        u2 = interp1(massattcoef.(material){:,1},massattcoef.(material){:,2},E); %mass attenuation coef. after reflection
        u2_ = interp1(massattcoef.(material){:,1},massattcoef.(material){:,3},E); % energy mass att. coef. after reflection
        C1 = interp1(Parameters{:,1},Parameters{:,2},E);
        C1_ = interp1(Parameters{:,1},Parameters{:,3},E);
        K = ((electronrad^2)/2)*((E_/E)^2)*((E/E_) + (E_/E) - (sind(scatangle))^2);
        a1 = (C1*K)/(u1+u2*(cosd(incidentangleEditField{1}.Value)/cosd(refangleEditField{1}.Value)));
        a1_ = (C1_)/(u1+u2_*(cosd(incidentangleEditField{1}.Value)/cosd(refangleEditField{1}.Value)));
        a = a1 + a1_;
        % [(uGym^2/MBqh)*MBq/m^2]*(m^2/m^2) + [uGy/h] = uGy/h
        MazeDoseRate(j) = (selectedSource.RAKR*activityEditField.Value*numberSourcesEditField.Value*F/mazedistEditField{1}.Value^2)*(a*1e-2*(mazeareaaEditField{1}.Value + mazeareabEditField{1}.Value)/(mazedistEditField{2}.Value^2)) + designLimitEditField{8}.Value*F;
        mazeData{j,1} = material;
        mazeData{j,2} = sprintf('%.2e uSv/h', MazeDoseRate(j));
        mazeData{j,3} = sprintf('NoMazeLeg');
    elseif cbx_twolegmaze.Value
        for i = 1:2
            scatangle{i} = 180 - (incidentangleEditField{i}.Value + refangleEditField{i}.Value);
        end
        E_ = E/(1+((E/0.511)*(1-cosd(scatangle{1}))));
        E1_ = E_/(1+((E_/0.511)*(1-cosd(scatangle{2}))));
        u1 = interp1(massattcoef.Air{:,1},massattcoef.Air{:,2},E);
        u2 = interp1(massattcoef.(material){:,1},massattcoef.(material){:,2},E); %mass attenuation coef. after reflection
        u2_ = interp1(massattcoef.(material){:,1},massattcoef.(material){:,3},E); % energy mass att. coef. after reflection
        u11 = interp1(massattcoef.Air{:,1},massattcoef.Air{:,2},E_);
        u22 = interp1(massattcoef.(material){:,1},massattcoef.(material){:,2},E_); %mass attenuation coef. after reflection
        u22_ = interp1(massattcoef.(material){:,1},massattcoef.(material){:,3},E_); % energy mass att. coef. after reflection
        C1 = interp1(Parameters{:,1},Parameters{:,2},E);
        C1_ = interp1(Parameters{:,1},Parameters{:,3},E);
        C2 = interp1(Parameters{:,1},Parameters{:,2},E_);
        C2_ = interp1(Parameters{:,1},Parameters{:,3},E_);
        K = ((electronrad^2)/2)*((E_/E)^2)*((E/E_) + (E_/E) - (sind(scatangle{1}))^2);
        K_ = ((electronrad^2)/2)*((E1_/E_)^2)*((E_/E1_) + (E1_/E_) - (sind(scatangle{2}))^2);
        a1 = (C1*K)/(u1+u2*(cosd(incidentangleEditField{1}.Value)/cosd(refangleEditField{1}.Value)));
        a1_ = (C1_)/(u11+u2_*(cosd(incidentangleEditField{1}.Value)/cosd(refangleEditField{1}.Value)));
        a = a1 + a1_;
        a2 = (C2*K_)/(u11+u22*(cosd(incidentangleEditField{2}.Value)/cosd(refangleEditField{2}.Value)));
        a2_ = (C2_)/(u11+u22_*(cosd(incidentangleEditField{2}.Value)/cosd(refangleEditField{2}.Value)));
        a_ = a2 + a2_;
        MazeDoseRate(j) = ((selectedSource.RAKR*activityEditField.Value*numberSourcesEditField.Value*F/mazedistEditField{1}.Value^2)*(a*1e-2*(mazeareaaEditField{1}.Value + mazeareabEditField{1}.Value)/(mazedistEditField{2}.Value^2))) + (designLimitEditField{9}.Value*F);
        if isnan(MazeDoseRate(j))
            MazeDoseRate(j) = 0;
        end
        MazeDoseRate1(j) = ((MazeDoseRate(j) - designLimitEditField{9}.Value*F)*((a_*1e-2*(mazeareaaEditField{2}.Value + mazeareabEditField{2}.Value))/((mazedistEditField{3}.Value^2)))) + (designLimitEditField{9}.Value*F);
        if isnan(MazeDoseRate1(j))
            MazeDoseRate1(j) = 0;
        end

        mazeData{j,1} = material;
        mazeData{j,2} = sprintf('%.2e uSv/h', MazeDoseRate(j));
        mazeData{j,3} = sprintf('%.2e uSv/h', MazeDoseRate1(j));
    else
        mazeData{j,1} = material;
        mazeData{j,2} = sprintf('NoMaze');
        mazeData{j,3} = sprintf('NoMaze');
    end
end

mazeInDoseRate = cell(1,2);
mazeDoseRate = cell(3,2);
for j = 1:length(materials)
    material = materials{j};
    if cbx_onelegmaze.Value
        mazeInDoseRate{1,1} = max(MazeDoseRate);
        mazeInDoseRate{1,2} = 0;
        mazeDoseRate{j,1} = mazeInDoseRate{1,1}*exp(-selectedSource.(material)*density.(material)*ceil(mazethickness(j,1)));
        mazeDoseRate{j,2} = 0;
        % Check if Activity is valid; if not, set to 0
        for i = 1:2
            if isnan(mazeInDoseRate{1,i}) || mazeInDoseRate{1,i} == Inf || mazeInDoseRate{1,i} == -Inf
                mazeInDoseRate{1,i} = 0;
            end
            if isnan(mazeDoseRate{j,i}) || mazeDoseRate{j,i} == Inf || mazeDoseRate{j,i} == -Inf
                mazeDoseRate{j,i} = 0;
            end
        end
        shieldData{1,28} = sprintf('%.2e uSv/h', mazeInDoseRate{1,1});
        shieldData{1,29} = 'NoLeg';
        shieldData{j+1,28} = sprintf('%.2e uSv/h', mazeDoseRate{j,1});
        shieldData{j+1,29} = 'NoLeg';
    elseif cbx_twolegmaze.Value
        for i = 1:2
            mazeInDoseRate{1,1} = max(MazeDoseRate);
            mazeInDoseRate{1,2} = max(MazeDoseRate1);
            mazeDoseRate{j,1} = mazeInDoseRate{1,1}*exp(-selectedSource.(material)*density.(material)*ceil(mazethickness(j,1)));
            mazeDoseRate{j,2} = mazeInDoseRate{1,2}*exp(-selectedSource.(material)*density.(material)*ceil(mazethickness_(j,1)));
            if isnan(mazeInDoseRate{1,i}) || mazeInDoseRate{1,i} == Inf || mazeInDoseRate{1,i} == -Inf
                mazeInDoseRate{1,i} = 0;
            end
            if isnan(mazeDoseRate{j,i}) || mazeDoseRate{j,i} == Inf || mazeDoseRate{j,i} == -Inf
                mazeDoseRate{j,i} = 0;
            end
        end
        shieldData{1,28} = sprintf('%.2e uSv/h', mazeInDoseRate{1,1});
        shieldData{1,29} = sprintf('%.2e uSv/h', mazeInDoseRate{1,2});
        shieldData{j+1,28} = sprintf('%.2e uSv/h', mazeDoseRate{j,1});
        shieldData{j+1,29} = sprintf('%.2e uSv/h', mazeDoseRate{j,2});
    else
        for i = 1:2
            mazeInDoseRate{1,i} = 0;
            mazeDoseRate{j,i} = 0;
        end
        shieldData{1,28} = 'NoLeg';
        shieldData{1,29} = 'NoLeg';
        shieldData{j+1,28} = 'NoLeg';
        shieldData{j+1,29} = 'NoLeg';
    end

    mixmazeDoseRate = zeros(1, 216);
    mixmazeCost = zeros(1, 216);
    mazeproportions = zeros(216, 3,1);
    mixmazeDoseRate_ = zeros(1,216);
    mixmazeCost_ = zeros(1, 216);
    mazeproportions_ = zeros(216, 3,1);

    % Initialize index for cells
    mazeindex = 1;
    mazeindex_ = 1;
    for p1 = 0:0.2:1
        for p2 = 0:0.2:1
            for p3 = 0:0.2:1
                % Calculate the mixed dose rate for each combination of proportions
                mixmazeDoseRate(mazeindex, 1) = mazeInDoseRate{1,1} * exp(-((selectedSource.(materials{1}) * density.(materials{1}) * ceil(mazethickness(1,1)) * p1) + (selectedSource.(materials{2}) * density.(materials{2}) * ceil(mazethickness(2,1)) * p2) + (selectedSource.(materials{3}) * density.(materials{3}) * ceil(mazethickness(3,1)) * p3)));
                % Calculate the mixed volume
                mazevolume1 = mazeareaEditField.Value * 1e6 * mazethickness(1,1) * p1;
                mazevolume2 = mazeareaEditField.Value * 1e6 * mazethickness(2,1) * p2;
                mazevolume3 = mazeareaEditField.Value * 1e6 * mazethickness(3,1) * p3;
                % Calculate the mixed cost
                mixmazeCost(mazeindex, 1) = (PriceEditField.(materials{1}).Value * density.(materials{1})*mazevolume1 + PriceEditField.(materials{2}).Value * density.(materials{2})*mazevolume2 + PriceEditField.(materials{3}).Value * density.(materials{3})*mazevolume3);
                % Store proportions for this combination
                mazeproportions(mazeindex, :,1) = [p1, p2, p3];

                % Calculate the mixed dose rate for each combination of proportions
                mixmazeDoseRate_(mazeindex_, 1) = mazeInDoseRate{1,2} * exp(-((selectedSource.(materials{1}) * density.(materials{1}) * ceil(mazethickness_(1,1)) * p1) + (selectedSource.(materials{2}) * density.(materials{2}) * ceil(mazethickness_(2,1)) * p2) + (selectedSource.(materials{3}) * density.(materials{3}) * ceil(mazethickness_(3,1)) * p3)));
                % Calculate the mixed volume
                mazevolume1_ = mazeareaEditField_.Value * 1e6 * mazethickness_(1,1) * p1;
                mazevolume2_ = mazeareaEditField_.Value * 1e6 * mazethickness_(2,1) * p2;
                mazevolume3_ = mazeareaEditField_.Value * 1e6 * mazethickness_(3,1) * p3;
                % Calculate the mixed cost
                mixmazeCost_(mazeindex_, 1) = (PriceEditField.(materials{1}).Value * density.(materials{1})*mazevolume1_ + PriceEditField.(materials{2}).Value * density.(materials{2})*mazevolume2_ + PriceEditField.(materials{3}).Value * density.(materials{3})*mazevolume3_);
                % Store proportions for this combination
                mazeproportions_(mazeindex_, :,1) = [p1, p2, p3];

                % Increment the index
                mazeindex = mazeindex + 1;
                if mazeindex > 216
                    break;  % Exit loop if index exceeds the limit
                end
                mazeindex_ = mazeindex_ + 1;
                if mazeindex_ > 216
                    break;
                end
            end
        end
    end

    % Reset bestDoseRate and bestCost for each iteration of i
    bestmazeDoseRate = inf;
    bestmazeCost = inf;
    bestmazeProportions = [0, 0, 0];  % Initialize best proportions for each i

    bestmazeDoseRate_ = inf;
    bestmazeCost_ = inf;
    bestmazeProportions_ = [0, 0, 0];  % Initialize best proportions for each i

    % Find the best mixture based on dose rate and cost for this i
    for mazeindex = 2:216
        if (mixmazeDoseRate(mazeindex, 1) <= 0.075) && (mixmazeCost(mazeindex, 1) <= bestmazeCost)
            bestmazeDoseRate = mixmazeDoseRate(mazeindex, 1);
            bestmazeCost = mixmazeCost(mazeindex, 1);
            bestmazeProportions = mazeproportions(mazeindex, :,1);  % Store the best proportions for this i
        end

        if (mixmazeDoseRate_(mazeindex, 1) <= 0.075) && (mixmazeCost_(mazeindex, 1) <= bestmazeCost_)
            bestmazeDoseRate_ = mixmazeDoseRate_(mazeindex, 1);
            bestmazeCost_ = mixmazeCost_(mazeindex, 1);
            bestmazeProportions_ = mazeproportions_(mazeindex, :,1);  % Store the best proportions for this i
        end
    end

    if isnan(bestmazeCost) || bestmazeCost == Inf || bestmazeCost == -Inf
        bestmazeCost = 0;
    end
    if isnan(bestmazeCost_) || bestmazeCost_ == Inf || bestmazeCost_ == -Inf
        bestmazeCost_ = 0;
    end

    % Store the best results for each iteration i
    if cbx_onelegmaze.Value
        tableData{4, 14} = sprintf('%.2f, %.2f, %.2f', bestmazeProportions(1), bestmazeProportions(2), bestmazeProportions(3));
        tableData{4, 15} = sprintf('€ %.2e', bestmazeCost);
        tableData{4, 16} = 'NoLeg';
        tableData{4, 17} = 'NoLeg';
        shieldData{5, 28} = sprintf('%.2e uSv/h', bestmazeDoseRate);
        shieldData{5, 29} = 'NoLeg';
    elseif cbx_twolegmaze.Value
        tableData{4, 14} = sprintf('%.2f, %.2f, %.2f', bestmazeProportions(1), bestmazeProportions(2), bestmazeProportions(3));
        tableData{4, 15} = sprintf('€ %.2e', bestmazeCost);
        tableData{4, 16} = sprintf('%.2f, %.2f, %.2f', bestmazeProportions_(1), bestmazeProportions_(2), bestmazeProportions_(3));
        tableData{4, 17} = sprintf('€ %.2e', bestmazeCost_);
        shieldData{5, 28} = sprintf('%.2e uSv/h', bestmazeDoseRate);
        shieldData{5, 29} = sprintf('%.2e uSv/h', bestmazeDoseRate_);
    else
        tableData{4, 14} = 'NoLeg';
        tableData{4, 15} = 'NoLeg';
        tableData{4, 16} = 'NoLeg';
        tableData{4, 17} = 'NoLeg';
        shieldData{5, 28} = 'NoLeg';
        shieldData{5, 29} = 'NoLeg';
    end

    s1 = uistyle('BackgroundColor','r');
    s2 = uistyle('BackgroundColor','y');
    s3 = uistyle('BackgroundColor','g');

    for i = 1:2
        if cbx_idr.Value
            if mazeInDoseRate{1,i} < 7.5 %uSv/h
                addStyle(shieldTable,s3,'cell',[1,i+27]);
            elseif mazeInDoseRate{1,i} == 7.5
                addStyle(shieldTable,s2,'cell',[1,i+27]);
            else
                addStyle(shieldTable,s1,'cell',[1,i+27]);
            end
        else
            if mazeInDoseRate{1,i} < 3 %uSv/h (weekly)
                addStyle(shieldTable,s3,'cell',[1,i+27]);
            elseif mazeInDoseRate{1,i} == 3
                addStyle(shieldTable,s2,'cell',[1,i+27]);
            else
                addStyle(shieldTable,s1,'cell',[1,i+27]);
            end
        end

        if mazeDoseRate{j,1} < designLimitEditField{7}.Value
            addStyle(shieldTable,s3,'cell',[j+1,28]);
            addStyle(shieldTable,s3,'cell',[5,28]);
        elseif mazeDoseRate{j,i} == designLimitEditField{7}.Value
            addStyle(shieldTable,s2,'cell',[j+1,28]);
            addStyle(shieldTable,s2,'cell',[5,28]);
        else
            addStyle(shieldTable,s1,'cell',[j+1,28]);
            addStyle(shieldTable,s1,'cell',[5,28]);
        end
        if mazeDoseRate{j,2} < designLimitEditField{8}.Value
            addStyle(shieldTable,s3,'cell',[j+1,29]);
            addStyle(shieldTable,s3,'cell',[5,29]);
        elseif mazeDoseRate{j,2} == designLimitEditField{8}.Value
            addStyle(shieldTable,s2,'cell',[j+1,29]);
            addStyle(shieldTable,s2,'cell',[5,29]);
        else
            addStyle(shieldTable,s1,'cell',[j+1,29]);
            addStyle(shieldTable,s1,'cell',[5,29]);
        end
    end
end

% Update the UITable with the new data
resultTable.Data = tableData;
shieldTable.Data = shieldData;
mazeTable.Data = mazeData;

end

%% Exporting Data
% Function to save table data to Excel
function saveToExcel(resultTable, shieldTable, mazeTable, mainFig)

% Open a dialog for the user to select a directory
folderName = uigetdir('', 'Select Folder to Save Excel File');

% Check if the user selected a folder (not canceled)
if folderName == 0
    return; % User canceled the operation
end

% Define the Excel filename
excelFileName = fullfile(folderName, 'BrachyBarrier.xlsx');

% Convert cell array to table format for export
exportTable = cell2table(resultTable.Data, 'VariableNames', resultTable.ColumnName);
exportTable1 = cell2table(shieldTable.Data, 'VariableNames', shieldTable.ColumnName);
exportTable2 = cell2table(mazeTable.Data, 'VariableNames', mazeTable.ColumnName);

% Write the table to the Excel file
writetable(exportTable, excelFileName, 'Sheet', 1, 'Range', 'A1');
writetable(exportTable1,excelFileName, 'Sheet', 1, 'Range', 'A10');
writetable(exportTable2,excelFileName, 'Sheet', 1, 'Range', 'A20');

% Alert user of successful save
uialert(mainFig, 'Data saved successfully!', 'Save Confirmation');

end
