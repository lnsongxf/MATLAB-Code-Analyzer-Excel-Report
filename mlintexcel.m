function mlintexcel(fn,excelfn)
% MLINTEXCEL    Generate a Microsoft Excel spreadsheet containing results of an m-lint report
%
% MLINTEXCEL Generates a Microsoft Excel spreadsheet containing the m-lint
% report of all files in the current directory.
%
% MLINTEXCEL(FILENAME) Generates the spreadsheet containing the m-lint
% report of file FILENAME
%
% MLINTEXCEL(DIRNAME) Generates a spreadsheet containing the m-lint
% report of all files in directory DIRNAME
%
% MLINTEXCEL(FILENAME,EXCELFN) Writes to Microsoft Excel file EXCELFN

% Michelle Hirsch
% mhirsch@mathworks.com
% Copyright 2005-2014 The MathWorks, Inc

if nargin==0
	fn = pwd;
end;


%% Build list of files
filetype = exist(fn);

if filetype == 2		% file found
	fn = {which(fn)};	% Convert to cell and move on
	
elseif filetype == 7	% Directory
	files = dir([fn filesep '*.m']);
	fn = {files(:).name}';	% Build list of files
	
else
	error('File not found');
end

%% Run M-Lint report
infostruct = mlint(fn,'-struct');

%% Break down some report results
% Number of issues
NIssues = cellfun(@length,infostruct);
NIssuesTotal = sum(NIssues);

% Remove ones with no issues
NoIssues = NIssues==0;
infostruct(NoIssues) = [];
fn(NoIssues) = [];

% Filenames
% List each file name once per mlint issue
fns = {};
linenumbers = [];
messages = {};
for ii=1:length(infostruct)
	fns = [fns; cellstr(repmat(fn{ii},length(infostruct{ii}),1))];
	messages = [messages;{infostruct{ii}.message}'];

	try
		linenumbers = [linenumbers;[infostruct{ii}.line]'];
	catch	% In some cases there are two line numbers (one is 0)
		is = infostruct{ii};		% Current file's structure
		lines = [];
		for jj=1:length(is)
			temp = is(jj).line;
			if length(temp)==2
				temp = temp(1);
			end
			lines = [lines;temp];
		end
		linenumbers = [linenumbers;lines];
	end
end

if isempty(fn)
	disp('Congratulations!  M-Lint generated no warnings.  No report created.')
	return
end



%% Create Excel File
% | FILENAME | LINE NUMBER | MESSAGE | FIXED |
if ~exist('excelfn','var')
	excelfn = [pwd filesep 'mlintreport.xls'];
else
	excelfn = which(excelfn);		% Full file name
end;

% Delete if exists
if exist(excelfn,'file')
	delete(excelfn)
end

h = waitbar(0,'Generating Excel File');

% Header
xlswrite(excelfn,{'Filename','Line #','Message','Fixed?'})
waitbar(.25,h)

% File Names
xlswrite(excelfn,fns,['A2:A' num2str(NIssuesTotal+1)])
waitbar(.5,h)

% Line Numbers
xlswrite(excelfn,linenumbers,['B2:B' num2str(NIssuesTotal+1)])
waitbar(.75,h)

% Messages
xlswrite(excelfn,messages,['C2:C' num2str(NIssuesTotal+1)])
close(h)

%% Format Excel File
if ispc		% Uses ActiveX, which is Windows-only
	% Make Headers bold
	ModifyExcelCell(excelfn,'Sheet1','A1:D1','bold');  % Make headers bold

	% Turn on autofilters
	ModifyExcelCell(excelfn,'Sheet1','A1:D1','AutoFilter');

	% AutoFit columns
	FixExcelColumns(excelfn,'Sheet1')  % AutoFit column widths

	%% Open file
	winopen(excelfn)
end



function FixExcelColumns(filename,sheetname)
% Auto Fit column widths in an Excel worksheet
% Lazy - assume filename is in current directory.  fix this.
Excel = actxserver('Excel.Application');
op = invoke(Excel.Workbooks, 'open', filename);

set(Excel, 'Visible', 0);

Sheets = Excel.ActiveWorkBook.Sheets;
target_sheet = get(Sheets, 'Item', sheetname);
invoke(target_sheet, 'Activate');

Activesheet = Excel.Activesheet;
Activesheet.cells.EntireColumn.AutoFit();
% Save and clean up
invoke(op, 'Save');
invoke(Excel, 'Quit');
delete(Excel)


function varargout = ModifyExcelCell(varargin)
% ModifyExcelCell(filename,sheetname,cellrange,property) modifies the
% property of a cell in an Excel spreadsheet.
%
% ModifyExcelCell(filename,sheetname,cellrange,property,val) modifies the
% property of a range of cells in an Excel spreadsheet with a specified value val.
%
% hExcel = ModifyExcelCell(filename,sheetname, ...) Returns handles to an open
% Excel spreadsheet. (hExcel.Excel, hExcel,op)
%
% ModifyExcelCell(hExcel,sheetname,cellrange,property) uses the already open
% Excel spreadsheet

% filename,sheetname,cellrange,property,varargin
% Modify Cell property
%  - AutoFilter
%  - enable html

%% Parse inputs / open file
if isstruct(varargin{1})
    handles = varargin{1};
    Excel = handles.Excel;
    op = handles.op;
else
    filename = varargin{1};
    Excel = actxserver('Excel.Application');
    op = invoke(Excel.Workbooks, 'open', filename);
end

[sheetname,cellrange,property] = varargin{2:4};

% Make specified cell active.  I should learn how to do this properly!
Sheets = Excel.ActiveWorkBook.Sheets;
target_sheet = get(Sheets, 'Item', sheetname);
invoke(target_sheet, 'Activate');
Activesheet = Excel.Activesheet;
% Range = Activesheet.cells.Range(cellrange,cellrange);     % Activate specified cell

% Excel.Visible = 1;


switch lower(property)
    case 'autofilter'
        % AutoFilter
        Excel.ActiveCell.AutoFilter;
    case 'hyperlink'
        link = varargin{5};
        % Hyperlink
        Activesheet.Hyperlinks.Add(Range,link);
    case 'bold'
        Activesheet.Range(cellrange).Font.Bold = 1;
    case 'size'
        fontsize = varargin{5};
%         Excel.ActiveCell.Font.Size = fontsize;
        Activesheet.Range(cellrange).Font.Size = fontsize;
end;

if nargout      % Return handles
    handles.Excel = Excel;
    handles.op = op;
    varargout{1} = handles;
else
    % Save and clean up
    invoke(op, 'Save');
    invoke(Excel, 'Quit');
    delete(Excel)
end;







