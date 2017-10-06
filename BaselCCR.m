classdef (Sealed) BaselCCR < BaselInterface
    %% Properties: Instance
    properties (Access = private)
        Collaterals
        CurrentHedgingSet
        CurrentNettingSet
        Handles
        Initialized
        NettingSets
        Result
        Trades
    end
    
    %% Constructor
    methods (Access = public)
        function this = BaselCCR()
            fig = findall(0,'Tag','BaselCCR');
            
            if (~isempty(fig))
                if (isvalid(fig))
                    figure(fig);
                    return;
                else
                    delete(fig);
                end
            end
            
            warning('off','all');
            
            com.mathworks.mwswing.MJUtilities.initJIDE();
            javaaddpath(fullfile(pwd(),'BaselTools.jar'));
            
            this.Initialized = false;
            
            gui = struct();
            gui.gui_Name = mfilename;
            gui.gui_Singleton = false;
            gui.gui_Callback = [];
            gui.gui_OpeningFcn = @this.Form_Load;
            gui.gui_OutputFcn = @this.Form_Output;
            gui.gui_LayoutFcn = [];
            
            gui_mainfcn(gui);
        end
    end
    
    %% Destructor
    methods (Access = private)
        function delete(this)
            if (~isempty(this.Handles))
                this.Handles.BaselCCR.CloseRequestFcn = '';
                delete(this.Handles.BaselCCR);
            end
            
            this.Collaterals = [];
            this.CurrentHedgingSet = [];
            this.CurrentNettingSet = [];
            this.Handles = [];
            this.Initialized = [];
            this.NettingSets = [];
            this.Result = [];
            this.Trades = [];
        end
    end
    
    %% Methods: Events
    methods (Access = private)
        function Form_Close(this,obj,evd) %#ok<INUSD>
            delete(this);
        end
        
        function Form_Load(this,obj,evd,han,varargin) %#ok<INUSL>
            obj.CloseRequestFcn = @this.Form_Close;
            
            han.DatasetButtonClear.Callback = @this.DatasetButtonClear_Clicked;
            han.DatasetButtonLoad.Callback = @this.DatasetButtonLoad_Clicked;
            han.ResultButtonCalculate.Callback = @this.ResultButtonCalculate_Clicked;
            han.ResultButtonReset.Callback = @this.ResultButtonReset_Clicked;
            
            han.TabGroup = uitabgroup('Parent',obj);
            han.TabGroup.Units = 'pixels';
            han.TabGroup.Position = [2 1 1024 768];
            han.IntroductionTab = uitab('Parent',han.TabGroup);
            han.IntroductionTab.Tag = 'IntroductionTab';
            han.IntroductionTab.Title = 'Introduction';
            han.IntroductionPanel.Parent = han.IntroductionTab;
            han.DatasetTab = uitab('Parent',han.TabGroup);
            han.DatasetTab.Tag = 'DatasetTab';
            han.DatasetTab.Title = 'Dataset';
            han.DatasetPanel.Parent = han.DatasetTab;
            han.ResultTab = uitab('Parent',han.TabGroup);
            han.ResultTab.Tag = 'ResultTab';
            han.ResultTab.Title = 'Result';
            han.ResultPanel.Parent = han.ResultTab;
            
            guidata(obj,han);
            this.Handles = guidata(obj);
            
            uistack(han.Blank,'top');
            
            scr = groot;
            scr.Units = 'pixels';
            scr_siz = scr.ScreenSize;
            obj.Position = [((scr_siz(3) - 1024) / 2) ((scr_siz(4) - 768) / 2) 1024 768];
        end
        
        function varargout = Form_Output(this,obj,evd,han) %#ok<INUSD>
            varargout = {};
            
            import('java.awt.*');
            
            bar = waitbar(0,'Initializing...','CloseRequestFcn','','WindowStyle','modal');
            frm = Frame.getFrames();
            frm(end).setAlwaysOnTop(true);
            
            try
                uistack(this.Handles.Blank,'bottom');
                
                this.SetupBox(this.Handles.IntroductionBox);
                
                this.SetupTable(this.Handles.DatasetTableNetting, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                this.SetupTable(this.Handles.DatasetTableCollaterals, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                this.SetupTable(this.Handles.DatasetTableTrades, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.SetupTable(this.Handles.ResultTableNetting, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                this.SetupTable(this.Handles.ResultTableHedging, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                this.SetupTable(this.Handles.ResultTableRisk, ...
                    'VerticalScrollbar', true);
                
                this.Initialized = true;
            catch e
                delete(bar);
                
                err = this.FormatException('The initialization process failed.',e);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                delete(this);
                
                return;
            end
            
            drawnow();
            pause(0.05);
            
            waitbar(1,bar);
            delete(bar);
        end
        
        function DatasetButtonClear_Clicked(this,obj,evd) %#ok<INUSD>
            obj.Enable = 'off';
            
            this.Handles.DatasetTextNetting.String = 'Netting Sets';
            this.SetupTable(this.Handles.DatasetTableNetting, ...
                'Sorting',           2, ...
                'VerticalScrollbar', true);
            
            this.Handles.DatasetTextCollaterals.String = 'Collaterals';
            this.SetupTable(this.Handles.DatasetTableCollaterals, ...
                'Sorting',           2, ...
                'VerticalScrollbar', true);
            
            this.Handles.DatasetTextTrades.String = 'Trades';
            this.SetupTable(this.Handles.DatasetTableTrades, ...
                'Sorting',           2, ...
                'VerticalScrollbar', true);
            
            this.Collaterals = [];
            this.NettingSets = [];
            this.Trades = [];
            
            this.Handles.DatasetCheckboxValidation.Enable = 'on';
            this.Handles.DatasetButtonLoad.Enable = 'on';

            this.SetupTable(this.Handles.ResultTableNetting, ...
                'Sorting',           2, ...
                'VerticalScrollbar', true);
            this.SetupTable(this.Handles.ResultTableHedging, ...
                'Sorting',           2, ...
                'VerticalScrollbar', true);
            this.SetupTable(this.Handles.ResultTableRisk, ...
                'VerticalScrollbar', true);
            
            this.CurrentHedgingSet = '';
            this.CurrentNettingSet = '';
            this.Result = [];
            
            this.Handles.ResultButtonCalculate.Enable = 'off';
            this.Handles.ResultButtonReset.Enable = 'off';
            this.Handles.ResultCheckboxOffset.Enable = 'off';
        end
        
        function DatasetButtonLoad_Clicked(this,obj,evd) %#ok<INUSD>
            obj.Enable = 'off';
            this.Handles.DatasetCheckboxValidation.Enable = 'off';
            
            [name,path] = uigetfile({'*.xls;*.xlsx','Excel Spreadsheets (*.xls;*.xlsx)'},'Load Dataset',[pwd() '\Datasets']);
            
            if (name == 0)
                obj.Enable = 'on';
                return;
            end
            
            import('java.awt.*');
            
            bar = waitbar(0,'Loading Dataset...','CloseRequestFcn','','WindowStyle','modal');
            frm = Frame.getFrames();
            frm(end).setAlwaysOnTop(true);
            
            try
                file = fullfile(path,name);
                [nss,trds,cols,err] = this.DatasetLoad(file);
            catch e
                err = this.FormatException('The dataset could not be loaded.',e);
            end
            
            if (~isempty(err))
                delete(bar);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                this.Handles.DatasetCheckboxValidation.Enable = 'on';
                obj.Enable = 'on';
                
                return;
            end
            
            waitbar(0.25,bar,'Validating Dataset...');
            
            try
                err = this.DatasetValidate(nss,trds,cols);
            catch e
                err = this.FormatException('The dataset could not be validated.',e);
            end
            
            if (~isempty(err))
                delete(bar);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                this.Handles.DatasetCheckboxValidation.Enable = 'on';
                obj.Enable = 'on';
                
                return;
            end
            
            waitbar(0.50,bar,'Analyzing Dataset...');
            
            try
                data = this.DatasetAnalyze(nss,cols,trds);
            catch e
                err = this.FormatException('The dataset could not be analyzed.',e);
            end
            
            if (~isempty(err))
                delete(bar);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                this.Handles.DatasetCheckboxValidation.Enable = 'on';
                obj.Enable = 'on';
                
                return;
            end
            
            waitbar(0.75,bar,'Updating Interface...');
            
            try
                this.Handles.DatasetTextNetting.String = ['Netting Sets (' num2str(size(data.Sets,1)) ')'];
                this.SetupTable(this.Handles.DatasetTableNetting, ...
                    'Data',              data.Sets, ...
                    'FillHeight',        false, ...
                    'Renderer',          {'RendererCcrDataNets' data.SetsMaximumID data.SetsMaximumThreshold data.SetsMaximumMTA data.SetsMaximumMPOR}, ...
                    'RowsHeight',        24, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.Handles.DatasetTextCollaterals.String = ['Collaterals (' num2str(size(data.Collaterals,1)) ')'];
                this.SetupTable(this.Handles.DatasetTableCollaterals, ...
                    'Data',              data.Collaterals, ...
                    'FillHeight',        false, ...
                    'Renderer',          {'RendererCcrDataCols' data.CollateralsMaximumID data.CollateralsMaximumLinkID data.CollateralsMaximumValue data.CollateralsMaximumMaturity}, ...
                    'RowsHeight',        24, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.Handles.DatasetTextTrades.String = ['Trades (' num2str(size(data.Trades,1)) ')'];
                this.SetupTable(this.Handles.DatasetTableTrades, ...
                    'Data',              data.Trades, ...
                    'FillHeight',        false, ...
                    'Renderer',          {'RendererCcrDataTrds' data.TradesMaximumID data.TradesMaximumNettingSetID data.TradesMaximumNotional data.TradesMaximumValue data.TradesMaximumStart data.TradesMaximumEnd data.TradesMaximumMaturity}, ...
                    'RowsHeight',        24, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.Handles.DatasetButtonClear.Enable = 'on';
                this.Handles.ResultCheckboxOffset.Enable = 'on';
                this.Handles.ResultButtonCalculate.Enable = 'on';
                
                this.Collaterals = cols;
                this.NettingSets = nss;
                this.Trades = trds;
            catch e
                err = this.FormatException('The interface could not be updated.',e);
            end
            
            if (~isempty(err))
                this.Handles.DatasetTextNetting.String = 'Netting Sets';
                this.SetupTable(this.Handles.DatasetTableNetting, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.Handles.DatasetTextCollaterals.String = 'Collaterals';
                this.SetupTable(this.Handles.DatasetTableCollaterals, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.Handles.DatasetTextTrades.String = 'Trades';
                this.SetupTable(this.Handles.DatasetTableTrades, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                delete(bar);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                this.Handles.DatasetCheckboxValidation.Enable = 'on';
                obj.Enable = 'on';
                
                return;
            end
            
            waitbar(1,bar);
            delete(bar);
        end
        
        function ResultButtonCalculate_Clicked(this,obj,evd) %#ok<INUSD>
            obj.Enable = 'off';
            this.Handles.ResultCheckboxOffset.Enable = 'off';
            
            import('java.awt.*');
            
            bar = waitbar(0,'Performing Calculations...','CloseRequestFcn','','WindowStyle','modal');
            frm = Frame.getFrames();
            frm(end).setAlwaysOnTop(true);
            
            err = '';
            
            try
                nss_len = height(this.NettingSets);
                nss_res = cell(nss_len,10);
                
                ead_tot = 0;
                
                for i = 1:nss_len
                    ns = this.NettingSets(i,:);
                    trds = this.Trades((this.Trades.NettingSetID == ns.ID),:);
                    cols = this.Collaterals((this.Collaterals.NettingSetID == ns.ID),:);
                    
                    if (isempty(trds))
                        nss_res(i,:) = {[num2str(ns.ID) 'N'] 0 NaN NaN NaN {} NaN NaN NaN {}};
                        continue;
                    end
                    
                    if (ns.Margined)
                        [rc_umar,pfe_umar,ead_umar,hs_umar] = this.CalculateEAD(ns,trds,cols,false);
                        [rc_mar,pfe_mar,ead_mar,hs_mar] = this.CalculateEAD(ns,trds,cols,true);
                        
                        if (ead_mar > ead_umar)
                            ead = ead_umar;
                        else
                            ead = ead_mar;
                        end
                        
                        nss_res(i,:) = {[num2str(ns.ID) 'N'] ead rc_umar pfe_umar ead_umar hs_umar rc_mar pfe_mar ead_mar hs_mar};
                    else
                        [rc,pfe,ead,hs] = this.CalculateEAD(ns,trds,cols,false);
                        nss_res(i,:) = {[num2str(ns.ID) 'N'] ead rc pfe ead hs NaN NaN NaN {}};
                    end
                    
                    ead_tot = ead_tot + ead;
                end
                
                nss_tab = nss_res(:,[1:5 7:9]);
                nss_vals = cell2mat(nss_tab(:,2:end));
                nss_max_id = max(cellfun(@(x)numel(x),nss_tab(:,1)));
                nss_max_ead = max(nss_vals(:,1));
                nss_max_rc_umar = this.FindConstrainedMaximum(nss_vals(:,2));
                nss_max_pfe_umar = this.FindConstrainedMaximum(nss_vals(:,3));
                nss_max_ead_umar = this.FindConstrainedMaximum(nss_vals(:,4));
                nss_max_rc_mar = this.FindConstrainedMaximum(nss_vals(:,5));
                nss_max_pfe_mar = this.FindConstrainedMaximum(nss_vals(:,6));
                nss_max_ead_mar = this.FindConstrainedMaximum(nss_vals(:,7));
            catch e
                err = this.FormatException('The calculations could not be performed.',e);
            end
            
            if (~isempty(err))
                delete(bar);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                this.Handles.ResultCheckboxOffset.Enable = 'on';
                obj.Enable = 'on';
                
                return;
            end
            
            waitbar(0.50,bar,'Updating Interface...');
            
            try
                import('baseltools.*');
                
                this.SetupTable(this.Handles.ResultTableNetting, ...
                    'Data',              nss_tab, ...
                    'FillHeight',        false, ...
                    'MouseClicked',      @this.ResultTableNetting_Clicked, ...
                    'MouseEntered',      @this.ResultTable_Entered, ...
                    'MouseExited',       @this.ResultTable_Exited, ...
                    'Renderer',          {'RendererCcrRsltNets' nss_max_id nss_max_ead nss_max_rc_umar nss_max_pfe_umar nss_max_ead_umar nss_max_rc_mar nss_max_pfe_mar nss_max_ead_mar}, ...
                    'RowsHeight',        24, ...
                    'Selection',         [false false true], ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                this.Handles.ResultTextEAD.String = ['Total EAD: ' char(Environment.FormatNumber(ead_tot,true,2))];
                
                this.SetupTable(this.Handles.ResultTableHedging, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.SetupTable(this.Handles.ResultTableRisk, ...
                    'VerticalScrollbar', true);

                this.Result = nss_res(:,[1 6 end]);
                this.CurrentNettingSet = '';
                this.CurrentHedgingSet = '';
                
                this.Handles.ResultButtonReset.Enable = 'on';
            catch e
                err = this.FormatException('The interface could not be updated.',e);
            end
            
            if (~isempty(err))
                this.SetupTable(this.Handles.ResultTableNetting, ...
                    'VerticalScrollbar', true);
                this.Handles.ResultTextEAD.String = 'Total EAD: 0,00';
                
                this.SetupTable(this.Handles.ResultTableHedging, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.SetupTable(this.Handles.ResultTableRisk, ...
                    'VerticalScrollbar', true);
                
                this.Result = [];
                this.CurrentNettingSet = '';
                this.CurrentHedgingSet = '';
                
                this.Handles.ResultButtonReset.Enable = 'off';
                
                delete(bar);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                this.Handles.ResultCheckboxOffset.Enable = 'on';
                obj.Enable = 'on';
                
                return;
            end
            
            waitbar(1,bar);
            delete(bar);
        end
        
        function ResultButtonReset_Clicked(this,obj,evd) %#ok<INUSD>
            obj.Enable = 'off';

            this.SetupTable(this.Handles.ResultTableNetting, ...
                'Sorting',           2, ...
                'VerticalScrollbar', true);
            this.SetupTable(this.Handles.ResultTableHedging, ...
                'Sorting',           2, ...
                'VerticalScrollbar', true);
            this.SetupTable(this.Handles.ResultTableRisk, ...
                'VerticalScrollbar', true);
            
            this.CurrentHedgingSet = '';
            this.CurrentNettingSet = '';
            this.Result = [];
            
            this.Handles.ResultButtonCalculate.Enable = 'on';
            this.Handles.ResultCheckboxOffset.Enable = 'on';
        end
        
        function ResultTable_Entered(this,obj,evd) %#ok<INUSD>
            this.Handles.BaselCCR.Pointer = 'hand';
        end
        
        function ResultTable_Exited(this,obj,evd) %#ok<INUSD>
            this.Handles.BaselCCR.Pointer = 'arrow';
        end
        
        function ResultTableHedging_Clicked(this,obj,evd)
            sel = obj.rowAtPoint(evd.getPoint());
            
            if (sel < 0)
                return;
            end
            
            hsid = obj.getValueAt(sel,0);
            
            if (strcmp(hsid,this.CurrentHedgingSet))
                return;
            end
            
            import('java.awt.*');
            
            bar = waitbar(0,'Displaying Risk Factors...','CloseRequestFcn','','WindowStyle','modal');
            frm = Frame.getFrames();
            frm(end).setAlwaysOnTop(true);
            
            err = '';
            
            try
                ns = this.Result(strcmp(this.Result(:,1),this.CurrentNettingSet),:);
                hs_umar = ns{2};
                
                if (isempty(hs_umar))
                    this.SetupTable(this.Handles.ResultTableRisk, ...
                        'VerticalScrollbar', true);
                else
                    rf_umar = hs_umar{strcmp(hs_umar(:,1),hsid),4};

                    if (isempty(rf_umar))
                        this.SetupTable(this.Handles.ResultTableRisk, ...
                            'VerticalScrollbar', true);
                    else
                        hs_mar = ns{3};
                        
                        if (isempty(hs_mar))
                            rf_mar = num2cell(nan(size(rf_umar)));
                            rf_mar(:,1) = rf_umar(:,1);
                        else
                            rf_mar = hs_mar{strcmp(hs_mar(:,1),hsid),4};
                        end

                        ref_vals_len = size(rf_umar,1);
                        ref_vals = cell((ref_vals_len * 2),6);

                        i_off = 1;

                        for i = 1:ref_vals_len
                            ref_vals(i_off,:) = {rf_umar{i,1} 'U' rf_umar{i,2:end}};
                            ref_vals(i_off+1,:) = {'' 'M' rf_mar{i,2:end}};

                            i_off = i_off + 2;
                        end

                        rf_max_eff_not = this.FindConstrainedLength(cell2mat(ref_vals(:,3)));
                        rf_max_addon = this.FindConstrainedLength(cell2mat(ref_vals(:,4)));
                        rf_max_comp_sys = this.FindConstrainedLength(cell2mat(ref_vals(:,5)));
                        rf_max_comp_idi = this.FindConstrainedLength(cell2mat(ref_vals(:,6)));

                        this.SetupTable(this.Handles.ResultTableRisk, ...
                            'Data',              ref_vals, ...
                            'FillHeight',        false, ...
                            'Renderer',          {'RendererCcrRsltRisk' hsid rf_max_eff_not rf_max_addon rf_max_comp_sys rf_max_comp_idi}, ...
                            'RowsHeight',        24, ...
                            'VerticalScrollbar', true);
                    end
                end

                this.CurrentHedgingSet = hsid;
            catch e
                err = this.FormatException('The risk factors could not be displayed.',e);
            end
            
            if (~isempty(err))
                this.SetupTable(this.Handles.ResultTableRisk, ...
                    'VerticalScrollbar', true);
                
                this.CurrentHedgingSet = hsid;
                
                delete(bar);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                return;
            end
            
            waitbar(1,bar);
            delete(bar);
        end
        
        function ResultTableNetting_Clicked(this,obj,evd)
            sel = obj.rowAtPoint(evd.getPoint());
            
            if (sel < 0)
                return;
            end
            
            nsid = obj.getValueAt(sel,0);
            
            if (strcmp(nsid,this.CurrentNettingSet))
                return;
            end
            
            import('java.awt.*');
            
            bar = waitbar(0,'Displaying Hedging Sets...','CloseRequestFcn','','WindowStyle','modal');
            frm = Frame.getFrames();
            frm(end).setAlwaysOnTop(true);
            
            err = '';
            
            try
                ns = this.Result(strcmp(this.Result(:,1),nsid),:);
                hs_umar = ns{2};
                hs_mar = ns{3};
                
                if (isempty(hs_umar))
                    this.SetupTable(this.Handles.ResultTableHedging, ...
                        'Sorting',           2, ...
                        'VerticalScrollbar', true);
                else
                    hs_vals_umar = cell2mat(hs_umar(:,2:3));
                    hs_max_eff_not_umar = this.FindConstrainedMaximum(hs_vals_umar(:,1));
                    hs_max_addon_umar = this.FindConstrainedMaximum(hs_vals_umar(:,2));
                    
                    if (isempty(hs_mar))
                        hs_mar = num2cell(nan(size(hs_umar)));
                        hs_mar(:,1) = hs_umar(:,1);
                    end
                    
                    hs_vals_mar = cell2mat(hs_mar(:,2:3));
                    hs_max_eff_not_mar = this.FindConstrainedMaximum(hs_vals_mar(:,1));
                    hs_max_addon_mar = this.FindConstrainedMaximum(hs_vals_mar(:,2));
                    
                    this.SetupTable(this.Handles.ResultTableHedging, ...
                        'Data',              [hs_umar(:,1:3) hs_mar(:,2:3)], ...
                        'FillHeight',        false, ...
                        'MouseClicked',      @this.ResultTableHedging_Clicked, ...
                        'MouseEntered',      @this.ResultTable_Entered, ...
                        'MouseExited',       @this.ResultTable_Exited, ...
                        'Renderer',          {'RendererCcrRsltHeds' hs_max_eff_not_umar hs_max_addon_umar hs_max_eff_not_mar hs_max_addon_mar}, ...
                        'RowsHeight',        24, ...
                        'Selection',         [false false true], ...
                        'Sorting',           2, ...
                        'VerticalScrollbar', true);
                end
                
                this.SetupTable(this.Handles.ResultTableRisk, ...
                    'VerticalScrollbar', true);
                
                this.CurrentNettingSet = nsid;
                this.CurrentHedgingSet = '';
            catch e
                err = this.FormatException('The hedging sets could not be displayed.',e);
            end
            
            if (~isempty(err))
                this.SetupTable(this.Handles.ResultTableHedging, ...
                    'Sorting',           2, ...
                    'VerticalScrollbar', true);
                
                this.SetupTable(this.Handles.ResultTableRisk, ...
                    'VerticalScrollbar', true);
                
                this.CurrentNettingSet = nsid;
                this.CurrentHedgingSet = '';
                
                delete(bar);
                
                dlg = errordlg(err,'Error','modal');
                uiwait(dlg);
                
                return;
            end
            
            waitbar(1,bar);
            delete(bar);
        end
    end

    %% Methods: Functions
    methods (Access = private)
        function [addon,hs] = CalculateAddonCommodities(this,trds,trds_mat_fac)
            if (isempty(trds))
                addon = 0;
                hs = {};

                return;
            end
            
            trds_len = height(trds);
            trds_par = cell(trds_len,3);
            
            for i = 1:trds_len
                trd = trds(i,:);
                trd_ref = char(trd.Reference);
                
                if (ismissing(trd.OptionPosition))
                    if (trd.Position == 'LONG')
                        trd_sup_del = 1;
                    else
                        trd_sup_del = -1;
                    end
                else
                    if (strcmp(trd_ref,'ELECTRICITY'))
                        trd_sup_vol = 1.5;
                    else
                        trd_sup_vol = 0.7;
                    end
                    
                    trd_sup_del = (log(trd.OptionPrice / trd.OptionStrike) + (0.5 * trd_sup_vol^2 * trd.OptionTime)) / (trd_sup_vol * sqrt(trd.OptionTime));
                    
                    if (trd.OptionPosition == 'LONG')
                        trd_sup_del = normcdf(trd_sup_del);
                    else
                        trd_sup_del = normcdf(-trd_sup_del);
                    end
                    
                    if (trd.Position == 'SHORT')
                        trd_sup_del = -1 * trd_sup_del;
                    end
                end
                
                if (isnan(trds_mat_fac))
                    trd_mat_fac = sqrt(min([1 trd.Maturity]));
                else
                    trd_mat_fac = trds_mat_fac;
                end
                
                trd_eff_not = trd.Notional * trd_sup_del * trd_mat_fac;
                
                trds_par(i,:) = {char(trd.Subclass) trd_ref trd_eff_not};
            end
            
            grp = findgroups(trds_par(:,1));
            hs = splitapply(@(x)this.CalculateAddonCommoditiesHedgingSet(x),trds_par,grp);
            hs = sortrows(hs,1);
            
            addon = 0;
            
            for i = 1:size(hs,1)
                addon = addon + hs{i,3};
            end
        end
        
        function [hs,rfs] = CalculateAddonCommoditiesHedgingSet(this,trds) %#ok<INUSL>
            hs_comp_idi = 0;
            hs_comp_sys = 0;
            hs_trds = trds(:,2:end);
            
            [grp,grp_id] = findgroups(hs_trds(:,1));
            grp_id_len = numel(grp_id);
            
            rfs = cell(grp_id_len,5);
            
            for i = 1:grp_id_len
                rf = hs_trds((grp == i),:);
                rf_name = grp_id{i};
                
                rf_eff_not = sum([rf{:,2}]);
                
                if (strcmp(rf_name,'ELECTRICITY'))
                    rf_addon = 0.4 * rf_eff_not;
                else
                    rf_addon = 0.18 * rf_eff_not;
                end
                
                rf_comp_idi = (1 - 0.16) * rf_addon^2;
                rf_comp_sys = 0.4 * rf_addon;
                
                hs_comp_idi = hs_comp_idi + rf_comp_idi;
                hs_comp_sys = hs_comp_sys + rf_comp_sys;
                
                rfs(i,:) = {rf_name rf_eff_not rf_addon rf_comp_sys rf_comp_idi};
            end
            
            hs = {['COMMODITIES - ' trds{1,1}] NaN sqrt(hs_comp_sys^2 + hs_comp_idi) rfs};
        end
        
        function [addon,hs] = CalculateAddonCredit(this,trds,trds_mat_fac) %#ok<INUSL>
            if (isempty(trds))
                addon = 0;
                hs = {};
                
                return;
            end
            
            trds_len = height(trds);
            trds_par = cell(trds_len,4);
            
            for i = 1:trds_len
                trd = trds(i,:);
                trd_ref = char(trd.Reference);
                
                if (trd.End < 1)
                    trd_sup_dur = sqrt(trd.End);
                else
                    trd_sup_dur = (exp(-0.05 * trd.Start) - exp(-0.05 * trd.End)) / 0.05;
                end
                
                trd_adj_not = trd.Notional * trd_sup_dur;
                
                if (~ismissing(trd.CDOAttachment))
                    trd_sup_del = 15 / ((1 + (14 * trd.CDOAttachment)) * (1 + (14 * trd.CDODetachment)));
                    
                    if (trd.Position == 'SHORT')
                        trd_sup_del = -1 * trd_sup_del;
                    end
                elseif (~ismissing(trd.OptionPosition))
                    if (trd.Class == 'CR_IDX')
                        trd_sup_del = (log(trd.OptionPrice / trd.OptionStrike) + (0.5 * 0.8^2 * trd.OptionTime)) / (0.8 * sqrt(trd.OptionTime));
                    else
                        trd_sup_del = (log(trd.OptionPrice / trd.OptionStrike) + (0.5 * trd.OptionTime)) / sqrt(trd.OptionTime);
                    end
                    
                    if (trd.Position == 'LONG')
                        if (trd.OptionPosition == 'LONG')
                            trd_sup_del = normcdf(trd_sup_del);
                        else
                            trd_sup_del = normcdf(-trd_sup_del);
                        end
                    else
                        if (trd.OptionPosition == 'LONG')
                            trd_sup_del = -normcdf(-trd_sup_del);
                        else
                            trd_sup_del = -normcdf(trd_sup_del);
                        end
                    end
                else
                    if (trd.Position == 'LONG')
                        trd_sup_del = 1;
                    else
                        trd_sup_del = -1;
                    end
                end
                
                if (isnan(trds_mat_fac))
                    trd_mat_fac = sqrt(min([1 trd.Maturity]));
                else
                    trd_mat_fac = trds_mat_fac;
                end
                
                trd_eff_not = trd_adj_not * trd_sup_del * trd_mat_fac;
                
                trds_par(i,:) = {trd.Class trd.Subclass trd_ref trd_eff_not};
            end
            
            comp_idi = 0;
            comp_sys = 0;
            
            [ref_uni,~,ref_uni_idx] = unique(trds_par(:,3));
            ref_uni_len = numel(ref_uni);
            
            rfs = [ref_uni num2cell(zeros(ref_uni_len,4))];
            
            for i = 1:ref_uni_len
                rf = trds_par((ref_uni_idx == i),:);
                rf_name = rf{1,3};
                rf_idx = strcmp(rfs(:,1),rf_name);
                
                switch (rf{1,1})
                    case 'CR_IDX'
                        rf_cor = 0.8;
                        
                        if (rf{1,2} == 'IG')
                            rf_sup_fac = 0.0038;
                        else
                            rf_sup_fac = 0.0106;
                        end
                        
                    case 'CR_SIN'
                        rf_cor = 0.5;
                        
                        switch (rf{1,2})
                            case {'AAA' 'AA'}
                                rf_sup_fac = 0.0038;
                            case 'A'
                                rf_sup_fac = 0.0042;
                            case 'BBB'
                                rf_sup_fac = 0.0054;
                            case 'B'
                                rf_sup_fac = 0.0160;
                            case 'CCC'
                                rf_sup_fac = 0.0600;
                            otherwise
                                rf_sup_fac = 0.0106;
                        end
                end
                
                rf_eff_not = sum([rf{:,4}]);
                rf_addon = rf_sup_fac * rf_eff_not;
                
                rf_comp_idi = (1 - rf_cor^2) * rf_addon^2;
                rf_comp_sys = rf_cor * rf_addon;
                
                comp_idi = comp_idi + rf_comp_idi;
                comp_sys = comp_sys + rf_comp_sys;
                
                rfs(rf_idx,:) = {rf_name rf_eff_not rf_addon rf_comp_sys rf_comp_idi};
            end
            
            addon = sqrt(comp_sys^2 + comp_idi);
            hs = {'CREDIT' NaN addon rfs};
        end
        
        function [addon,hs] = CalculateAddonEquity(this,trds,trds_mat_fac) %#ok<INUSL>
            if (isempty(trds))
                addon = 0;
                hs = {};
                
                return;
            end
            
            trds_len = height(trds);
            trds_par = cell(trds_len,3);
            
            for i = 1:trds_len
                trd = trds(i,:);
                trd_cls = char(trd.Class);
                trd_ref = char(trd.Reference);
                
                if (trd.End < 1)
                    trd_sup_dur = sqrt(trd.End);
                else
                    trd_sup_dur = (exp(-0.05 * trd.Start) - exp(-0.05 * trd.End)) / 0.05;
                end
                
                trd_adj_not = trd.Notional * trd_sup_dur;
                
                if (ismissing(trd.OptionPosition))
                    if (trd.Position == 'LONG')
                        trd_sup_del = 1;
                    else
                        trd_sup_del = -1;
                    end
                else
                    if (trd.Class == 'EQ_IDX')
                        trd_sup_del = (log(trd.OptionPrice / trd.OptionStrike) + (0.5 * 0.75^2 * trd.OptionTime)) / (0.75 * sqrt(trd.OptionTime));
                    else
                        trd_sup_del = (log(trd.OptionPrice / trd.OptionStrike) + (0.5 * 1.2^2 * trd.OptionTime)) / (1.2 * sqrt(trd.OptionTime));
                    end
                    
                    if (trd.Position == 'LONG')
                        if (trd.OptionPosition == 'LONG')
                            trd_sup_del = normcdf(trd_sup_del);
                        else
                            trd_sup_del = normcdf(-trd_sup_del);
                        end
                    else
                        if (trd.OptionPosition == 'LONG')
                            trd_sup_del = -normcdf(-trd_sup_del);
                        else
                            trd_sup_del = -normcdf(trd_sup_del);
                        end
                    end
                end
                
                if (isnan(trds_mat_fac))
                    trd_mat_fac = sqrt(min([1 trd.Maturity]));
                else
                    trd_mat_fac = trds_mat_fac;
                end
                
                trd_eff_not = trd_adj_not * trd_sup_del * trd_mat_fac;
                
                trds_par(i,:) = {trd_cls trd_ref trd_eff_not};
            end
            
            comp_idi = 0;
            comp_sys = 0;
            
            [ref_uni,~,ref_uni_idx] = unique(trds_par(:,2));
            ref_uni_len = numel(ref_uni);
            
            rfs = [ref_uni num2cell(zeros(ref_uni_len,4))];
            
            for i = 1:ref_uni_len
                rf = trds_par((ref_uni_idx == i),:);
                rf_name = rf{1,2};
                rf_idx = strcmp(rfs(:,1),rf_name);
                
                if (rf{1,1} == 'EQ_IDX')
                    rf_cor = 0.8;
                    rf_sup_fac = 0.2;
                else
                    rf_cor = 0.5;
                    rf_sup_fac = 0.32;
                end
                
                rf_eff_not = sum([rf{:,3}]);
                rf_addon = rf_sup_fac * rf_eff_not;
                
                rf_comp_idi = (1 - rf_cor^2) * rf_addon^2;
                rf_comp_sys = rf_cor * rf_addon;
                
                comp_idi = comp_idi + rf_comp_idi;
                comp_sys = comp_sys + rf_comp_sys;
                
                rfs(rf_idx,:) = {rf_name rf_eff_not rf_addon rf_comp_sys rf_comp_idi};
            end
            
            addon = sqrt(comp_sys^2 + comp_idi);
            hs = {'EQUITY' NaN addon rfs};
        end
        
        function [addon,hs] = CalculateAddonForex(this,trds,trds_mat_fac) %#ok<INUSL>
            if (isempty(trds))
                addon = 0;
                hs = {};
                
                return;
            end
            
            trds_len = height(trds);
            trds_par = cell(trds_len,2);
            
            for i = 1:trds_len
                trd = trds(i,:);
                trd_ccys = sort({char(trd.PayerCurrency) char(trd.ReceiverCurrency)});
                trd_hs = [trd_ccys{1} '/' trd_ccys{2}];
                
                if (ismissing(trd.OptionPosition))
                    if (trd.Position == 'LONG')
                        trd_sup_del = 1;
                    else
                        trd_sup_del = -1;
                    end
                else
                    trd_sup_del = (log(trd.OptionPrice / trd.OptionStrike) + (0.5 * 0.15^2 * trd.OptionTime)) / (0.15 * sqrt(trd.OptionTime));
                    
                    if (trd.Position == 'LONG')
                        if (trd.OptionPosition == 'LONG')
                            trd_sup_del = normcdf(trd_sup_del);
                        else
                            trd_sup_del = normcdf(-trd_sup_del);
                        end
                    else
                        if (trd.OptionPosition == 'LONG')
                            trd_sup_del = -normcdf(-trd_sup_del);
                        else
                            trd_sup_del = -normcdf(trd_sup_del);
                        end
                    end
                end
                
                if (isnan(trds_mat_fac))
                    trd_mat_fac = sqrt(min([1 trd.Maturity]));
                else
                    trd_mat_fac = trds_mat_fac;
                end
                
                trd_eff_not = trd.Notional * trd_sup_del * trd_mat_fac;
                
                trds_par(i,:) = {trd_hs trd_eff_not};
            end
            
            [hs_uni,~,hs_uni_idx] = unique(trds_par(:,1));
            hs_uni_len = numel(hs_uni);
            
            addon = 0;
            hs = cell(hs_uni_len,4);
            
            for i = 1:hs_uni_len
                hs_trds_par = trds_par((hs_uni_idx == i),:);
                
                hs_eff_not = sum([hs_trds_par{:,2}]);
                hs_addon = 0.04 * abs(hs_eff_not);
                
                addon = addon + hs_addon;
                
                hs(i,1:end) = {['FOREIGN EXCHANGE - ' hs_trds_par{1,1}] hs_eff_not hs_addon {}};
            end
            
            hs = sortrows(hs,1);
        end
        
        function [addon,hs] = CalculateAddonInterestRate(this,trds,trds_mat_fac)
            if (isempty(trds))
                addon = 0;
                hs = {};
                
                return;
            end
            
            trds_len = height(trds);
            trds_par = cell(trds_len,3);
            
            for i = 1:trds_len
                trd = trds(i,:);
                trd_pay_leg = char(trd.PayerLeg);
                trd_rec_leg = char(trd.ReceiverLeg);
                
                if (~strcmp(trd_pay_leg,'FIXED') && ~strcmp(trd_rec_leg,'FIXED'))
                    trd_pay_spl = strsplit(trd_pay_leg,'-');
                    
                    if (endsWith(trd_pay_spl(2),'ON'))
                        trd_pay_freq = 1;
                    else
                        trd_pay_tenor = trd_pay_spl{2};
                        
                        switch (trd_pay_tenor(end))
                            case 'W'
                                trd_pay_freq = 5;
                            case 'M'
                                trd_pay_freq = 21;
                            case 'Y'
                                trd_pay_freq = 252;
                            otherwise
                                trd_pay_freq = 1;
                        end
                        
                        trd_pay_freq = trd_pay_freq * str2double(trd_pay_tenor(1:end-1));
                    end
                    
                    trd_rec_spl = strsplit(char(trd_rec_leg),'-');
                    
                    if (endsWith(trd_rec_spl(2),'ON'))
                        trd_rec_freq = 1;
                    else
                        trd_rec_tenor = trd_rec_spl{2};
                        
                        switch (trd_rec_tenor(end))
                            case 'W'
                                trd_rec_freq = 5;
                            case 'M'
                                trd_rec_freq = 21;
                            case 'Y'
                                trd_rec_freq = 252;
                            otherwise
                                trd_rec_freq = 1;
                        end
                        
                        trd_rec_freq = trd_rec_freq * str2double(trd_rec_tenor(1:end-1));
                    end
                    
                    trd_freq = {trd_pay_leg trd_pay_freq; trd_rec_leg trd_rec_freq};
                    trd_freq = sortrows(trd_freq,2);
                    trd_freq_1 = trd_freq{1};
                    trd_freq_2 = trd_freq{2};
                    
                    if (strcmp(trd_freq_1,trd_freq_2))
                        trd_type = ['B ' char(trd.PayerCurrency) ' ' strrep(trd_freq_1,'-','')];
                    else
                        trd_type = ['B ' char(trd.PayerCurrency) ' ' strrep(trd_freq_1,'-','') '/' strrep(trd_freq_2,'-','')];
                    end
                else
                    if (~strcmp(trd_pay_leg,'FIXED') && endsWith(trd_pay_leg,'-VOL'))
                        trd_type = ['V ' char(trd.PayerCurrency) ' ' strrep(trd_pay_leg,'-VOL','')];
                    elseif (~strcmp(trd_rec_leg,'FIXED') && endsWith(trd_rec_leg,{'-VAR' '-VOL'}))
                        trd_type = ['V ' char(trd.PayerCurrency) ' ' strrep(trd_rec_leg,'-VOL','')];
                    else
                        trd_type = char(trd.PayerCurrency);
                    end
                end
                
                if (trd.Maturity < 1)
                    trd_buc = 1;
                elseif (trd.Maturity <= 5)
                    trd_buc = 2;
                else
                    trd_buc = 3;
                end
                
                if (trd.End < 1)
                    trd_sup_dur = sqrt(trd.End);
                else
                    trd_sup_dur = (exp(-0.05 * trd.Start) - exp(-0.05 * trd.End)) / 0.05;
                end
                
                trd_adj_not = trd.Notional * trd_sup_dur;
                
                if (ismissing(trd.OptionPosition))
                    if (strcmp(trd_pay_leg,'FIXED'))
                        trd_sup_del = 1;
                    else
                        trd_sup_del = -1;
                    end
                else
                    trd_sup_del = (log(trd.OptionPrice / trd.OptionStrike) + (0.5 * 0.5^2 * trd.OptionTime)) / (0.5 * sqrt(trd.OptionTime));
                    
                    if (strcmp(trd_pay_leg,'FIXED'))
                        if (trd.OptionPosition == 'LONG')
                            trd_sup_del = normcdf(trd_sup_del);
                        else
                            trd_sup_del = normcdf(-trd_sup_del);
                        end
                    else
                        if (trd.OptionPosition == 'LONG')
                            trd_sup_del = -normcdf(-trd_sup_del);
                        else
                            trd_sup_del = -normcdf(trd_sup_del);
                        end
                    end
                end
                
                if (isnan(trds_mat_fac))
                    trd_mat_fac = sqrt(min([1 trd.Maturity]));
                else
                    trd_mat_fac = trds_mat_fac;
                end
                
                trd_eff_not = trd_adj_not * trd_sup_del * trd_mat_fac;
                
                trds_par(i,:) = {trd_type trd_buc trd_eff_not};
            end
            
            grp = findgroups(trds_par(:,1));
            hs = splitapply(@(x)this.CalculateAddonInterestRateHedgingSet(x),trds_par,grp);
            hs = sortrows(hs,[1 2 3]);
            
            for i = 1:size(hs,1)
                hs_pre = hs{i,1};
                
                if (strcmp(hs_pre,''))
                    hs{i,1} = ['INTEREST RATE - ' hs{i,2}];
                else
                    hs{i,1} = ['INTEREST RATE - ' hs{i,2} ' ' hs{i,1} ' ' hs{i,3}];
                end
            end
            
            hs(:,2:3) = [];
            
            addon = 0;
            
            for i = 1:size(hs,1)
                addon = addon + hs{i,3};
            end
        end
        
        function [hs,rfs] = CalculateAddonInterestRateHedgingSet(this,trds)
            hs_name = trds{1,1};
            hs_trds = trds(:,2:end);
            
            [grp,grp_id] = findgroups([hs_trds{:,1}]);
            
            rfs = {'MATURITY BUCKET 1' 0 NaN NaN NaN; 'MATURITY BUCKET 2' 0 NaN NaN NaN; 'MATURITY BUCKET 3' 0 NaN NaN NaN;};
            
            for i = 1:numel(grp_id)
                rf = hs_trds((grp == i),:);
                rf_buc = grp_id(i);
                
                rf_eff_not = sum([rf{:,2}]);
                
                rfs{rf_buc,2} = rf_eff_not;
            end
            
            if (this.Handles.ResultCheckboxOffset.Value == 0)
                hs_eff_not = abs(rfs{1,2}) + abs(rfs{2,2}) + abs(rfs{3,2});
            else
                rfs_eff_not_1 = rfs{1,2};
                rfs_eff_not_2 = rfs{2,2};
                rfs_eff_not_3 = rfs{3,2};
                rfs_corr_12 = 1.4 * rfs_eff_not_1 * rfs_eff_not_2;
                rfs_corr_13 = 0.6 * rfs_eff_not_1 * rfs_eff_not_3;
                rfs_corr_23 = 1.4 * rfs_eff_not_2 * rfs_eff_not_3;
                
                hs_eff_not = sqrt(rfs_eff_not_1^2 + rfs_eff_not_2^2 + rfs_eff_not_3^2 + rfs_corr_12 + rfs_corr_23 + rfs_corr_13);
            end
            
            if (startsWith(hs_name,'B '))
                hs_name_spl = strsplit(hs_name,' ');
                hs_name_pre = 'BASIS';
                hs_name_ccy = hs_name_spl{2};
                hs_name_tar = hs_name_spl{3};
                
                hs_sup_fac = 0.0025;
            elseif (startsWith(hs_name,'B '))
                hs_name_spl = strsplit(hs_name,' ');
                hs_name_pre = 'VOLATILITY';
                hs_name_ccy = hs_name_spl{2};
                hs_name_tar = hs_name_spl{3};
                
                hs_sup_fac = 0.025;
            else
                hs_name_pre = '';
                hs_name_ccy = hs_name;
                hs_name_tar = '';
                
                hs_sup_fac = 0.005;
            end
            
            hs = {hs_name_pre hs_name_ccy hs_name_tar hs_eff_not (hs_sup_fac * hs_eff_not) rfs};
        end
        
        function c = CalculateCollateral(this,cols,col_mat_fac) %#ok<INUSL>
            c = 0;
            
            for i = 1:height(cols)
                col = cols(i,:);
                
                switch (col.Type)
                    case 'BOND_SOVEREIGN'
                        switch (char(col.Parameter))
                            case {'AAA' 'AA' 'A-1'}
                                if (col.Maturity <= 1)
                                    col_sup_hc = 0.005;
                                elseif (col.Maturity <= 5)
                                    col_sup_hc = 0.02;
                                else
                                    col_sup_hc = 0.04;
                                end
                            case {'A' 'BBB' 'A-2' 'A-3'}
                                if (col.Maturity <= 1)
                                    col_sup_hc = 0.01;
                                elseif (col.Maturity <= 5)
                                    col_sup_hc = 0.03;
                                else
                                    col_sup_hc = 0.06;
                                end
                            case 'BB'
                                col_sup_hc = 0.15;
                        end
                    case 'BOND_OTHER'
                        switch (char(col.Parameter))
                            case {'AAA' 'AA' 'A-1'}
                                if (col.Maturity <= 1)
                                    col_sup_hc = 0.01;
                                elseif (col.Maturity <= 5)
                                    col_sup_hc = 0.04;
                                else
                                    col_sup_hc = 0.08;
                                end
                            case {'A' 'BBB' 'A-2' 'A-3'}
                                if (col.Maturity <= 1)
                                    col_sup_hc = 0.02;
                                elseif (col.Maturity <= 5)
                                    col_sup_hc = 0.06;
                                else
                                    col_sup_hc = 0.12;
                                end
                        end
                    case 'CASH'
                        col_sup_hc = 0;
                    case {'EQUITY_MAIN' 'GOLD'}
                        col_sup_hc = 0.15;
                    case 'EQUITY_OTHER'
                        col_sup_hc = 0.25;
                    case 'OTHER'
                        col_sup_hc = str2double(char(col.Parameter));
                end
                
                col_hc = 5 * col_sup_hc;
                
                if (~isnan(col_mat_fac))
                    col_hc = col_hc * col_mat_fac;
                end
                
                if (col.Value < 0)
                    col_val = min([0 (col.Value + abs(col_hc * col.Value))]);
                else
                    col_val = col.Value + (col_hc * col.Value);
                end
                
                c = c + col_val;
            end
        end
        
        function [rc,pfe,ead,hs] = CalculateEAD(this,ns,trds,cols,mar)
            if (mar)
                cols_mat_fac = sqrt(double(ns.MPOR) / 250);
                trds_mat_fac = 1.5 * sqrt(max([10 double(ns.MPOR)]) / 250);
                rc3 = ns.Threshold + ns.MTA + this.CalculateNICA(cols);
            else
                cols_mat_fac = NaN;
                trds_mat_fac = NaN;
                rc3 = NaN;
            end
            
            v = sum(trds.Value);
            c = this.CalculateCollateral(cols,cols_mat_fac);
            vc = v + c;
            rc = nanmax([0 vc rc3]);
            
            trd_co = trds((trds.Class == 'CO'),:);
            [addon_co,hs_co] = this.CalculateAddonCommodities(trd_co,trds_mat_fac);
            
            trd_cr = trds(((trds.Class == 'CR_IDX') | (trds.Class == 'CR_SIN')),:);
            [addon_cr,hs_cr] = this.CalculateAddonCredit(trd_cr,trds_mat_fac);
            
            trd_eq = trds(((trds.Class == 'EQ_IDX') | (trds.Class == 'EQ_SIN')),:);
            [addon_eq,hs_eq] = this.CalculateAddonEquity(trd_eq,trds_mat_fac);
            
            trd_fx = trds((trds.Class == 'FX'),:);
            [addon_fx,hs_fx] = this.CalculateAddonForex(trd_fx,trds_mat_fac);
            
            trd_ir = trds((trds.Class == 'IR'),:);
            [addon_ir,hs_ir] = this.CalculateAddonInterestRate(trd_ir,trds_mat_fac);
            
            addon = addon_co + addon_cr + addon_eq + addon_fx + addon_ir;
            hs = [hs_co; hs_cr; hs_eq; hs_fx; hs_ir];
            
            if (vc < 0)
                mul = min([1 (0.05 + (0.95 * exp(vc / (1.9 * addon))))]);
            else
                mul = 1;
            end
            
            pfe = mul * addon;
            
            ead = 1.4 * (rc + pfe);
        end
        
        function nica = CalculateNICA(this,cols) %#ok<INUSL>
            nica = 0;
            
            for i = 1:height(cols)
                col = cols(i,:);
                
                if (col.Margin == 'ICA')
                    nica = nica + col.Value;
                end
            end
        end
        
        function data = DatasetAnalyze(this,nss,cols,trds)
            nss_len = height(nss);
            nss_tab = cell(nss_len,5);
            
            for i = 1:nss_len
                ns = nss(i,:);
                
                nss_tab(i,1:2) = {ns.ID ns.Margined};
                
                if (ns.Margined)
                    nss_tab(i,3:5) = {ns.Threshold ns.MTA ns.MPOR};
                else
                    nss_tab(i,3:5) = {NaN NaN NaN};
                end
            end
            
            cols_len = height(cols);
            cols_tab = cell(cols_len,7);
            
            for i = 1:cols_len
                col = cols(i,:);
                
                cols_tab{i,1} = col.ID;
                
                if (col.NettingSetID > 0)
                    cols_tab{i,2} = [num2str(col.NettingSetID) 'N'];
                else
                    cols_tab{i,2} = [num2str(col.TradeID) 'T'];
                end
                
                switch (col.Type)
                    case 'BOND_SOVEREIGN'
                        cols_tab{i,3} = 'BOND_SOV';
                    case 'BOND_OTHER'
                        cols_tab{i,3} = 'BOND_OTH';
                    case 'CASH'
                        cols_tab{i,3} = 'CASH';
                    case 'EQUITY_MAIN'
                        cols_tab{i,3} = 'EQUI_SOV';
                    case 'EQUITY_OTHER'
                        cols_tab{i,3} = 'EQUI_OTH';
                    case 'GOLD'
                        cols_tab{i,3} = 'GOLD';
                    case 'OTHER'
                        cols_tab{i,3} = 'OTHER';
                end
                
                cols_tab(i,4:5) = {char(col.Margin) col.Value};
                
                if (ismember(col.Type,categorical({'BOND_SOVEREIGN' 'BOND_OTHER' 'OTHER'})))
                    cols_tab{i,6} = char(col.Parameter);
                else
                    cols_tab{i,6} = '';
                end
                
                if (ismember(col.Type,categorical({'BOND_SOVEREIGN' 'BOND_OTHER'})))
                    cols_tab{i,7} = col.Maturity;
                else
                    cols_tab{i,7} = NaN;
                end
            end
            
            trds_len = height(trds);
            trds_tab = cell(trds_len,11);
            
            for i = 1:trds_len
                trd = trds(i,:);
                
                trds_tab(i,1:2) = {trd.ID trd.NettingSetID};
                
                if (~ismissing(trd.Subclass))
                    cls_sub = char(trd.Subclass);
                    cls_sub_len = min([3 length(cls_sub)]);
                    cls_sub_code = cls_sub(1:cls_sub_len);
                    
                    trds_tab{i,3} = [char(trd.Class) ' / ' cls_sub_code];
                else
                    trds_tab{i,3} = char(trd.Class);
                end
                
                trds_tab(i,4:8) = {trd.Notional trd.Value trd.Start trd.End trd.Maturity};
                
                switch (trd.Class)
                    case 'FX'
                        trds_tab{i,9} = [char(trd.Position) ' | ' char(trd.PayerCurrency) '/' char(trd.ReceiverCurrency)];
                    case 'IR'
                        trds_tab{i,9} = [char(trd.PayerCurrency) '/' char(trd.PayerLeg) ' + ' char(trd.ReceiverCurrency) '/' char(trd.ReceiverLeg)];
                    otherwise
                        trds_tab{i,9} = [char(trd.Position) ' | ' char(trd.Reference)];
                end
                
                if (ismember(trd.Class,categorical({'CR_IDX' 'CR_SIN'})))
                    if (~ismissing(trd.CDOAttachment))
                        trds_tab{i,10} = 'YES';
                    else
                        trds_tab{i,10} = 'NO';
                    end
                else
                    trds_tab{i,10} = 'N/A';
                end
                
                if (~ismissing(trd.OptionPosition))
                    trds_tab{i,11} = 'YES';
                else
                    trds_tab{i,11} = 'NO';
                end
            end
            
            data = struct();
            data.Collaterals = cols_tab;
            data.CollateralsMaximumID = max(cols.ID);
            data.CollateralsMaximumLinkID = max(cellfun(@(x)numel(x),cols_tab(:,2)));
            data.CollateralsMaximumValue = this.FindConstrainedLength(cols.Value);
            data.CollateralsMaximumMaturity = max([0; cols.Maturity]);
            data.Sets = nss_tab;
            data.SetsMaximumID = max(nss.ID);
            data.SetsMaximumThreshold = max(nss.Threshold);
            data.SetsMaximumMTA = max(nss.MTA);
            data.SetsMaximumMPOR = max(nss.MPOR);
            data.Trades = trds_tab;
            data.TradesMaximumID = max(trds.ID);
            data.TradesMaximumNettingSetID = max(trds.NettingSetID);
            data.TradesMaximumNotional = max(trds.Notional);
            data.TradesMaximumValue = this.FindConstrainedLength(trds.Value);
            data.TradesMaximumStart = max(trds.Start);
            data.TradesMaximumEnd = max(trds.End);
            data.TradesMaximumMaturity = max(trds.Maturity);
        end
        
        function [nss,trds,cols,err] = DatasetLoad(this,file) %#ok<INUSL>
            nss = [];
            trds = [];
            cols = [];
            err = '';
            
            if (exist(file,'file') == 0)
                err = 'The dataset file does not exist.';
                return;
            end
            
            [file_sta,file_shts,file_fmt] = xlsfinfo(file);
            
            if (isempty(file_sta) || ~strcmp(file_fmt,'xlOpenXMLWorkbook'))
                err = 'The dataset file is not a valid Excel spreadsheet.';
                return;
            end
            
            if (numel(file_shts) ~= 3)
                err = 'The dataset must contain three sheets.';
                return;
            end
            
            opts = detectImportOptions(file,'Sheet',1);
            opts = setvartype(opts,{'uint32' 'logical' 'double' 'double' 'uint32'});
            opts = setvaropts(opts,'Margined','FalseSymbols',{'NO'},'TrueSymbols',{'YES'});
            nss = readtable(file,opts);
            
            opts = detectImportOptions(file,'Sheet',2);
            opts = setvartype(opts,{'uint32' 'uint32' 'categorical' 'categorical' 'double' 'double' 'double' 'double' 'double' 'categorical' 'char' 'char' 'char' 'char' 'char' 'double' 'double' 'categorical' 'double' 'double' 'double'});
            opts = setvaropts(opts,'Class','Categories',{'CO' 'CR_IDX' 'CR_SIN' 'EQ_IDX' 'EQ_SIN' 'FX' 'IR'});
            opts = setvaropts(opts,'Subclass','Categories',{'OTHER' 'AGRICULTURAL' 'ENERGY' 'METAL' 'IG' 'SG' 'AAA' 'AA' 'A' 'BBB' 'BB' 'B' 'CCC'});
            opts = setvaropts(opts,'Position','Categories',{'LONG' 'SHORT'});
            opts = setvaropts(opts,'OptionPosition','Categories',{'LONG' 'SHORT'});
            trds = readtable(file,opts);
            
            opts = detectImportOptions(file,'Sheet',3);
            opts = setvartype(opts,{'uint32' 'uint32' 'uint32' 'categorical' 'categorical' 'double' 'char' 'double'});
            opts = setvaropts(opts,'Type','Categories',{'BOND_SOVEREIGN' 'BOND_OTHER' 'CASH' 'EQUITY_MAIN' 'EQUITY_OTHER' 'GOLD' 'OTHER'});
            opts = setvaropts(opts,'Margin','Categories',{'ICA' 'VM'});
            cols = readtable(file,opts);
        end
        
        function err = DatasetValidate(this,nss,trds,cols)
            stg = this.Handles.DatasetCheckboxValidation.Value == 1;
            
            err = this.DatasetValidateNettingSets(stg,nss,trds);
            
            if (~isempty(err))
                return;
            end
            
            err = this.DatasetValidateTrades(stg,nss,trds);
            
            if (~isempty(err))
                return;
            end
            
            err = this.DatasetValidateCollaterals(stg,nss,trds,cols);
        end
        
        function err = DatasetValidateCollaterals(this,stg,nss,trds,cols) %#ok<INUSL>
            err = '';
            
            if (~isequal(strtrim(cols.Properties.VariableNames),{'ID' 'NettingSetID' 'TradeID' 'Type' 'Margin' 'Value' 'Parameter' 'Maturity'}))
                err = 'The ''Colleterals'' table is invalid because of wrong columns count, order and/or name.';
                return;
            end
            
            if (isempty(cols))
                return;
            end
            
            id_inv = (cols.ID == 0);
            
            if (any(id_inv))
                rows = (1:height(cols))';
                ref = rows(id_inv) + 1;
                err = char(strcat("In the 'Collaterals' table, the 'ID' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            [~,id_idx] = unique(cols.ID);
            id_dup = setdiff(1:numel(cols.ID),id_idx) + 1;
            
            if (~isempty(id_dup))
                err = char(strcat("In the 'Collaterals' table, the 'ID' column contains duplicate values (rows: ",strjoin(string(id_dup),', '),")."));
                return;
            end
            
            nsid_inv = (cols.NettingSetID == 0) | ((cols.NettingSetID > 0) & ~ismember(cols.NettingSetID,nss.ID));
            tid_inv = (cols.TradeID == 0) | ((cols.TradeID > 0) & ~ismember(cols.TradeID,trds.ID));
            link_inv = (nsid_inv & tid_inv) | (~nsid_inv & ~tid_inv);
            
            if (any(link_inv))
                rows = (1:height(cols))';
                ref = rows(link_inv) + 1;
                err = char(strcat("In the 'Collaterals' table, entries must either specify a valid 'NettingSetID' or a valid 'TradeID' (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            trds_inv = ismember(trds.ID,cols.TradeID) & (trds.NettingSetID > 0);
            
            if (any(trds_inv))
                rows = (1:height(cols))';
                ref = rows(cols(ismember(cols.ID,trds(trds_inv,:).ID),:)) + 1;
                err = char(strcat("In the 'Collaterals' table, entries can be linked to trades only when the latter don't belong to a netting set (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            type_inv = ismissing(cols.Type);
            
            if (any(type_inv))
                rows = (1:height(cols))';
                ref = rows(type_inv) + 1;
                err = char(strcat("In the 'Collaterals' table, the 'Type' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            mar_inv = ismissing(cols.Margin);
            
            if (any(mar_inv))
                rows = (1:height(cols))';
                ref = rows(mar_inv) + 1;
                err = char(strcat("In the 'Collaterals' table, the 'Margin' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            val_inv = ~isreal(cols.Value) | ~isfinite(cols.Value);
            
            if (any(val_inv))
                rows = (1:height(cols))';
                ref = rows(val_inv) + 1;
                err = char(strcat("In the 'Collaterals' table, the 'Value' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            par_o = str2double(cols.Parameter);
            par_mis = ismissing(cols.Parameter);
            par_inv_bs = (cols.Type == 'BOND_SOVEREIGN') & (par_mis | ~ismember(cols.Parameter,{'AAA','AA','A','BBB' 'BB' 'A-1' 'A-2' 'A-3'}));
            par_inv_bo = (cols.Type == 'BOND_OTHER') & (par_mis | ~ismember(cols.Parameter,{'AAA','AA','A','BBB' 'A-1' 'A-2' 'A-3'}));
            par_inv_o = (cols.Type == 'OTHER') & (~isreal(par_o) | ~isfinite(par_o) | (par_o < 0) | (par_o > 1));
            par_inv = par_inv_bs | par_inv_bo | par_inv_o;
            
            if (any(par_inv))
                rows = (1:height(cols))';
                ref = rows(par_inv) + 1;
                err = char(strcat("In the 'Collaterals' table, the 'Parameter' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                par_inv = ~ismember(cols.Type,categorical({'BOND_SOVEREIGN' 'BOND_OTHER' 'OTHER'})) & ~par_mis;
                
                if (any(par_inv))
                    rows = (1:height(cols))';
                    ref = rows(par_inv) + 1;
                    err = char(strcat("In the 'Collaterals' table, only bond or undefined type entries should specify a 'Parameter' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            cols_mat = ismember(cols.Type,categorical({'BOND_SOVEREIGN' 'BOND_OTHER'}));
            mat_inv = cols_mat & (~isreal(cols.Maturity) | ~isfinite(cols.Maturity) | (cols.Maturity <= 0));
            
            if (any(mat_inv))
                rows = (1:height(cols))';
                ref = rows(mat_inv) + 1;
                err = char(strcat("In the 'Collaterals' table, the 'Maturity' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                mat_inv = ~cols_mat & ~par_mis;
                
                if (any(mat_inv))
                    rows = (1:height(cols))';
                    ref = rows(mat_inv) + 1;
                    err = char(strcat("In the 'Collaterals' table, only bond entries should specify a 'Maturity' value (rows: ",strjoin(string(ref),', '),")."));
                end
            end
        end
        
        function err = DatasetValidateNettingSets(this,stg,nss,trds) %#ok<INUSL>
            err = '';
            
            if (~isequal(strtrim(nss.Properties.VariableNames),{'ID' 'Margined' 'Threshold' 'MTA' 'MPOR'}))
                err = 'The ''Netting Sets'' table is invalid because of wrong columns count, order and/or name.';
                return;
            end
            
            if (isempty(nss))
                return;
            end
            
            id_inv = (nss.ID == 0);
            
            if (any(id_inv))
                rows = (1:height(nss))';
                ref = rows(id_inv) + 1;
                err = char(strcat("In the 'Netting Sets' table, the 'ID' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            [~,id_idx] = unique(nss.ID);
            id_dup = setdiff(1:numel(nss.ID),id_idx) + 1;
            
            if (~isempty(id_dup))
                err = char(strcat("In the 'Netting Sets' table, the 'ID' column contains duplicate values (rows: ",strjoin(string(id_dup),', '),")."));
                return;
            end
            
            thr_inv = nss.Margined & (~isreal(nss.Threshold) | ~isfinite(nss.Threshold) | (nss.Threshold < 0));
            
            if (any(thr_inv))
                rows = (1:height(nss))';
                ref = rows(thr_inv) + 1;
                err = char(strcat("In the 'Netting Sets' table, the 'Threshold' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                thr_inv = ~nss.Margined & ~ismissing(nss.Threshold);
                
                if (any(thr_inv))
                    rows = (1:height(nss))';
                    ref = rows(thr_inv) + 1;
                    err = char(strcat("In the 'Netting Sets' table, unmargined entries should not specify a 'Threshold' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            mta_inv = nss.Margined & (~isreal(nss.MTA) | ~isfinite(nss.MTA) | (nss.MTA < 0));
            
            if (any(mta_inv))
                rows = (1:height(nss))';
                ref = rows(mta_inv) + 1;
                err = char(strcat("In the 'Netting Sets' table, the 'MTA' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                mta_inv = ~nss.Margined & ~ismissing(nss.MTA);
                
                if (any(mta_inv))
                    rows = (1:height(nss))';
                    ref = rows(mta_inv) + 1;
                    err = char(strcat("In the 'Netting Sets' table, unmargined entries should not specify an 'MTA' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            mpor_inv = nss.Margined & (~isreal(nss.MPOR) | ~isfinite(nss.MPOR) | (nss.MPOR < 10));
            
            if (any(mpor_inv))
                rows = (1:height(nss))';
                ref = rows(mpor_inv) + 1;
                err = char(strcat("In the 'Netting Sets' table, the 'MPOR' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                mpor_inv = ~nss.Margined & ~ismissing(nss.MPOR);
                
                if (any(mpor_inv))
                    rows = (1:height(nss))';
                    ref = rows(mpor_inv) + 1;
                    err = char(strcat("In the 'Netting Sets' table, unmargined entries should not specify an 'MPOR' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            if (stg)
                nss_unu = ~ismember(nss.ID,trds.NettingSetID);
                
                if (any(nss_unu))
                    rows = (1:height(nss))';
                    ref = rows(nss_unu) + 1;
                    err = char(strcat("The 'Netting Sets' table contains unused entries (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
        end
        
        function err = DatasetValidateTrades(this,stg,nss,trds)
            err = '';
            
            if (isempty(trds))
                err = 'The ''Trades'' table is empty.';
                return;
            end
            
            if (~isequal(strtrim(trds.Properties.VariableNames),{'ID' 'NettingSetID' 'Class' 'Subclass' 'Notional' 'Value' 'Start' 'End' 'Maturity' 'Position' 'Reference' 'PayerCurrency' 'ReceiverCurrency' 'PayerLeg' 'ReceiverLeg' 'CDOAttachment' 'CDODetachment' 'OptionPosition' 'OptionPrice' 'OptionStrike' 'OptionTime'}))
                err = 'The ''Trades'' table is invalid because of wrong columns count, order and/or name.';
                return;
            end
            
            id_inv = (trds.ID == 0);
            
            if (any(id_inv))
                rows = (1:height(trds))';
                ref = rows(id_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'ID' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            [~,id_idx] = unique(trds.ID);
            id_dup = setdiff(1:numel(trds.ID),id_idx) + 1;
            
            if (~isempty(id_dup))
                err = char(strcat("In the 'Trades' table, the 'ID' column contains duplicate values (rows: ",strjoin(string(id_dup),', '),")."));
                return;
            end
            
            nsid_inv = (trds.NettingSetID > 0) & ~ismember(trds.NettingSetID,nss.ID);
            
            if (any(nsid_inv))
                rows = (1:height(trds))';
                ref = rows(nsid_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'NettingSetID' column contains invalid values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            cls_inv = ismissing(trds.Class);
            
            if (any(cls_inv))
                rows = (1:height(trds))';
                ref = rows(cls_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'Class' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            trds_co = trds.Class == 'CO';
            trds_cr_idx = trds.Class == 'CR_IDX';
            trds_cr_sin = trds.Class == 'CR_SIN';
            trds_cr = trds_cr_idx | trds_cr_sin;
            trds_eq_idx = trds.Class == 'EQ_IDX';
            trds_eq_sin = trds.Class == 'EQ_SIN';
            trds_eq = trds_eq_idx | trds_eq_sin;
            trds_fx = trds.Class == 'FX';
            trds_ir = trds.Class == 'IR';
            trds_leg = trds_fx | trds_ir;
            
            cls_sub_inv_co = trds_co & ~ismember(trds.Subclass,categorical({'AGRICULTURAL','ENERGY','METAL','OTHER'}));
            cls_sub_inv_cr_idx = trds_cr_idx & ~ismember(trds.Subclass,categorical({'IG','SG'}));
            cls_sub_inv_cr_sin = trds_cr_sin & ~ismember(trds.Subclass,categorical({'AAA','AA','A','BBB','BB','B','CCC','OTHER'}));
            cls_sub_inv = cls_sub_inv_co | cls_sub_inv_cr_idx | cls_sub_inv_cr_sin;
            
            if (any(cls_sub_inv))
                rows = (1:height(trds))';
                ref = rows(cls_sub_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'Subclass' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                cls_sub_inv = (trds_eq | trds_leg) & ~ismissing(trds.Subclass);
                
                if (any(cls_sub_inv))
                    rows = (1:height(trds))';
                    ref = rows(cls_sub_inv) + 1;
                    err = char(strcat("In the 'Trades' table, only commodity and credit entries should specify a 'Subclass' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            not_inv = ~isreal(trds.Notional) | ~isfinite(trds.Notional) | (trds.Notional <= 0);
            
            if (any(not_inv))
                rows = (1:height(trds))';
                ref = rows(not_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'Notional' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            val_inv = ~isreal(trds.Value) | ~isfinite(trds.Value);
            
            if (any(val_inv))
                rows = (1:height(trds))';
                ref = rows(val_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'Value' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            sta_inv = ~isreal(trds.Start) | ~isfinite(trds.Start) | (trds.Start < 0) | (trds.Start >= trds.End);
            
            if (any(sta_inv))
                rows = (1:height(trds))';
                ref = rows(sta_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'Start' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            end_inv = ~isreal(trds.End) | ~isfinite(trds.End) | (trds.End <= 0);
            
            if (any(end_inv))
                rows = (1:height(trds))';
                ref = rows(end_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'End' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            mat_inv = ~isreal(trds.Maturity) | ~isfinite(trds.Maturity) | (trds.Maturity <= 0);
            
            if (any(mat_inv))
                rows = (1:height(trds))';
                ref = rows(mat_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'Maturity' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            pos_mis = ismissing(trds.Position);
            pos_inv = ~trds_ir & pos_mis;
            
            if (any(pos_inv))
                rows = (1:height(trds))';
                ref = rows(pos_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'Position' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                pos_inv = trds_ir & ~pos_mis;
                
                if (any(pos_inv))
                    rows = (1:height(trds))';
                    ref = rows(pos_inv) + 1;
                    err = char(strcat("In the 'Trades' table, interest rate entries should not specify a 'Position' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            ref_mis = ismissing(trds.Reference);
            ref_inv_co = trds_co & (trds.Subclass == 'ENERGY') & ~ismember(trds.Reference,{'COAL' 'CRUDE OIL' 'ELECTRICITY' 'NATURAL GAS'});
            
            if (stg)
                ref_inv_eq_sin = trds_eq_sin & ~this.IsValidISIN(trds.Reference);
            else
                ref_inv_eq_sin = trds_eq_sin & cellfun(@isempty,regexp(trds.Reference,'^[A-Z]{2}[0-9A-Z]{9}[0-9]{1}$'));
            end
            
            ref_inv = (~trds_leg & ref_mis) | ref_inv_co | ref_inv_eq_sin;
            
            if (any(ref_inv))
                rows = (1:height(trds))';
                ref = rows(ref_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'Reference' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                ref_inv = trds_leg & ~ref_mis;
                
                if (any(ref_inv))
                    rows = (1:height(trds))';
                    ref = rows(ref_inv) + 1;
                    err = char(strcat("In the 'Trades' table, foreign exchange and interest rate entries should not specify a 'Reference' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            if (stg)
                pc_inv = trds_leg & ~this.IsValidCurrency(trds.PayerCurrency);
            else
                pc_inv = trds_leg & cellfun(@isempty,regexp(trds.PayerCurrency,'^[A-Z]{3}$'));
            end
            
            if (any(pc_inv))
                rows = (1:height(trds))';
                ref = rows(pc_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'PayerCurrency' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                pc_inv = ~trds_leg & ~ismissing(trds.PayerCurrency);
                
                if (any(pc_inv))
                    rows = (1:height(trds))';
                    ref = rows(pc_inv) + 1;
                    err = char(strcat("In the 'Trades' table, only foreign exchange and interest rate entries should specify a 'PayerCurrency' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            if (stg)
                rc_inv = trds_leg & ~this.IsValidCurrency(trds.ReceiverCurrency);
            else
                rc_inv = trds_leg & cellfun(@isempty,regexp(trds.ReceiverCurrency,'^[A-Z]{3}$'));
            end
            
            if (any(rc_inv))
                rows = (1:height(trds))';
                ref = rows(rc_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'ReceiverCurrency' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                rc_inv = ~trds_leg & ~ismissing(trds.ReceiverCurrency);
                
                if (any(rc_inv))
                    rows = (1:height(trds))';
                    ref = rows(rc_inv) + 1;
                    err = char(strcat("In the 'Trades' table, only foreign exchange and interest rate entries should specify a 'ReceiverCurrency' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            ccy_equ = strcmp(trds.PayerCurrency,trds.ReceiverCurrency);
            prc_inv = (trds_fx & ccy_equ) | (trds_ir & ~ccy_equ);
            
            if (any(prc_inv))
                rows = (1:height(trds))';
                ref = rows(prc_inv) + 1;
                err = char(strcat("In 'Trades' table, foreign exchange entries must have different 'PayerCurrency' and 'ReceiverCurrency' values whereas interest rate entries must have identical 'PayerCurrency' and 'ReceiverCurrency' values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            pl_mis = ismissing(trds.PayerLeg);
            pl_inv = trds_ir & (pl_mis | (~strcmp(trds.PayerLeg,'FIXED') & cellfun(@isempty,regexp(trds.PayerLeg,'^[A-Z]+-(VOL|[0-9]{1,2}(D|W|M|Y))$'))));
            
            if (any(pl_inv))
                rows = (1:height(trds))';
                ref = rows(pl_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'PayerLeg' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                pl_inv = ~trds_ir & ~pl_mis;
                
                if (any(pl_inv))
                    rows = (1:height(trds))';
                    ref = rows(pl_inv) + 1;
                    err = char(strcat("In the 'Trades' table, only interest rate entries should specify a 'PayerLeg' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            rl_mis = ismissing(trds.ReceiverLeg);
            rl_inv = trds_ir & (rl_mis | (~strcmp(trds.ReceiverLeg,'FIXED') & cellfun(@isempty,regexp(trds.ReceiverLeg,'^[A-Z]+-(VOL|[0-9]{1,2}(D|W|M|Y))$'))));
            
            if (any(rl_inv))
                rows = (1:height(trds))';
                ref = rows(rl_inv) + 1;
                err = char(strcat("In the 'Trades' table, the 'ReceiverLeg' column contains invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                rl_inv = ~trds_ir & ~rl_mis;
                
                if (any(rl_inv))
                    rows = (1:height(trds))';
                    ref = rows(rl_inv) + 1;
                    err = char(strcat("In the 'Trades' table, only interest rate entries should specify a 'ReceiverLeg' value (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            cdo_att_mis = ismissing(trds.CDOAttachment);
            cdo_det_mis = ismissing(trds.CDODetachment);
            cdo_mis = sum([cdo_att_mis cdo_det_mis],2);
            
            cdo_att_inv = ~isreal(trds.CDOAttachment) | ~isfinite(trds.CDOAttachment) | (trds.CDOAttachment < 0) | (trds.CDOAttachment >= trds.CDODetachment);
            cdo_det_inv = ~isreal(trds.CDODetachment) | ~isfinite(trds.CDODetachment) | (trds.CDODetachment <= 0) | (trds.CDODetachment > 1);
            cdo_inv = trds_cr & ((cdo_mis == 1) | ((cdo_mis == 0) & (cdo_att_inv | cdo_det_inv)));
            
            if (any(cdo_inv))
                rows = (1:height(trds))';
                ref = rows(cdo_inv) + 1;
                err = char(strcat("The 'Trades' table contains credit entries linked to CDOs with invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            if (stg)
                cdo_inv = ~trds_cr & (cdo_mis < 2);
                
                if (any(cdo_inv))
                    rows = (1:height(trds))';
                    ref = rows(cdo_inv) + 1;
                    err = char(strcat("In the 'Trades' table, only credit entries can be linked to CDOs (rows: ",strjoin(string(ref),', '),")."));
                    return;
                end
            end
            
            opt_pos_mis = ismissing(trds.OptionPosition);
            opt_pri_mis = ismissing(trds.OptionPrice);
            opt_stk_mis = ismissing(trds.OptionStrike);
            opt_time_mis = ismissing(trds.OptionTime);
            opt_mis = sum([opt_pos_mis opt_pri_mis opt_stk_mis opt_time_mis],2);
            
            opt_pri_inv = ~isreal(trds.OptionPrice) | ~isfinite(trds.OptionPrice) | (trds.OptionPrice <= 0);
            opt_stk_inv = ~isreal(trds.OptionStrike) | ~isfinite(trds.OptionStrike) | (trds.OptionStrike <= 0);
            opt_time_inv = ~isreal(trds.OptionTime) | ~isfinite(trds.OptionTime) | (trds.OptionTime <= 0);
            opt_inv = ((opt_mis > 0) & (opt_mis < 4)) | ((opt_mis == 0) & (opt_pri_inv | opt_stk_inv | opt_time_inv));
            
            if (any(opt_inv))
                rows = (1:height(trds))';
                ref = rows(opt_inv) + 1;
                err = char(strcat("The 'Trades' table contains entries linked to options with invalid or missing values (rows: ",strjoin(string(ref),', '),")."));
                return;
            end
            
            cdo_opt_inv = trds_cr & (cdo_mis == 0) & (opt_mis == 0);
            
            if (any(cdo_opt_inv))
                rows = (1:height(trds))';
                ref = rows(cdo_opt_inv) + 1;
                err = char(strcat("In the 'Trades' table, credit entries cannot be linked to both a CDO and an option at the same time (rows: ",strjoin(string(ref),', '),")."));
            end
        end
        
        function res = FindConstrainedLength(this,vec) %#ok<INUSL>
            import('baseltools.*');
            
            vec_len = length(vec);
            len = zeros(vec_len,1);
            
            for i = 1:vec_len
                vec_val = vec(i);
                
                if (isnan(vec_val))
                    len(i) = 3;
                else
                    len(i) = length(Environment.FormatCurrencyGrouped.format(vec_val));
                end
            end
            
            res = max(len);
        end
        
        function res = FindConstrainedMaximum(this,vec) %#ok<INUSL>
            if (any(isnan(vec)))
                res = nanmax([3; vec]);
            else
                res = max(vec);
            end
        end
        
        function res = IsValidCurrency(this,ccy) %#ok<INUSL>
            persistent pat;
            
            if (isempty(pat))
                pat = strcat('^AED|AFN|ALL|AMD|ANG|AOA|ARS|AUD|AWG|AZN|BAM|BBD|BDT|BGN|BHD|BIF|BMD|BND|', ...
                    'BOB|BRL|BSD|BTN|BWP|BYR|BZD|CAD|CDF|CHF|CLP|CNY|COP|CRC|CUC|CUP|CVE|CZK|', ...
                    'DJF|DKK|DOP|DZD|EGP|ERN|ETB|EUR|FJD|FKP|GBP|GEL|GGP|GHS|GIP|GMD|GNF|GTQ|', ...
                    'GYD|HKD|HNL|HRK|HTG|HUF|IDR|ILS|IMP|INR|IQD|IRR|ISK|JEP|JMD|JOD|JPY|KES|', ...
                    'KGS|KHR|KMF|KPW|KRW|KWD|KYD|KZT|LAK|LBP|LKR|LRD|LSL|LYD|MAD|MDL|MGA|MKD|', ...
                    'MMK|MNT|MOP|MRO|MUR|MVR|MWK|MXN|MYR|MZN|NAD|NGN|NIO|NOK|NPR|NZD|OMR|PAB|', ...
                    'PEN|PGK|PHP|PKR|PLN|PYG|QAR|RON|RSD|RUB|RWF|SAR|SBD|SCR|SDG|SEK|SGD|SHP|', ...
                    'SLL|SOS|SPL|SRD|STD|SVC|SYP|SZL|THB|TJS|TMT|TND|TOP|TRY|TTD|TVD|TWD|TZS|', ...
                    'UAH|UGX|USD|UYU|UZS|VEF|VND|VUV|WST|XAF|XCD|XDR|XOF|XPF|YER|ZAR|ZMW|ZWD$');
            end
            
            res = ~cellfun(@isempty,regexp(ccy,pat));
        end
        
        function res = IsValidISIN(this,isin) %#ok<INUSL>
            persistent pat;
            
            if (isempty(pat))
                pat = strcat('^(AD|AE|AF|AG|AI|AL|AM|AO|AQ|AR|AS|AT|AU|AW|AX|AZ|BA|BB|BD|BE|BF|BG|BH|BI|BJ|', ...
                    'BL|BM|BN|BO|BQ|BR|BS|BT|BV|BW|BY|BZ|CA|CC|CD|CF|CG|CH|CI|CK|CL|CM|CN|CO|CR|', ...
                    'CU|CV|CW|CX|CY|CZ|DE|DJ|DK|DM|DO|DZ|EC|EE|EG|EH|ER|ES|ET|FI|FJ|FK|FM|FO|FR|', ...
                    'GA|GB|GD|GE|GF|GG|GH|GI|GL|GM|GN|GP|GQ|GR|GS|GT|GU|GW|GY|HK|HM|HN|HR|HT|HU|', ...
                    'ID|IE|IL|IM|IN|IO|IQ|IR|IS|IT|JE|JM|JO|JP|KE|KG|KH|KI|KM|KN|KP|KR|KW|KY|KZ|', ...
                    'LA|LB|LC|LI|LK|LR|LS|LT|LU|LV|LY|MA|MC|MD|ME|MF|MG|MH|MK|ML|MM|MN|MO|MP|MQ|', ...
                    'MR|MS|MT|MU|MV|MW|MX|MY|MZ|NA|NC|NE|NF|NG|NI|NL|NO|NP|NR|NU|NZ|OM|PA|PE|PF|', ...
                    'PG|PH|PK|PL|PM|PN|PR|PS|PT|PW|PY|QA|RE|RO|RS|RU|RW|SA|SB|SC|SD|SE|SG|SH|SI|', ...
                    'SJ|SK|SL|SM|SN|SO|SR|SS|ST|SV|SX|SY|SZ|TC|TD|TF|TG|TH|TJ|TK|TL|TM|TN|TO|TR|', ...
                    'TT|TV|TW|TZ|UA|UG|UM|US|UY|UZ|VA|VC|VE|VG|VI|VN|VU|WF|WS|XS|YE|YT|ZA|ZM|ZW)', ...
                    '[0-9A-Z]{9}[0-9]{1}$');
            end
            
            isin_len = numel(isin);
            res = false(isin_len,1);
            
            for i = 1:isin_len
                isin_curr = isin{i};
                
                if (isempty(isin_curr))
                    continue;
                end
                
                rm = regexp(isin_curr,pat,'match');
                
                if (isempty(rm))
                    continue;
                end
                
                code = reshape(isin_curr(1:end-1),1,[])';
                
                vals = zeros(11,1);
                vals_off = 0;
                
                for j = 1:11
                    code_curr = code(j);
                    
                    vals_off = vals_off + 1;
                    
                    if (isstrprop(code_curr,'digit'))
                        vals(vals_off) = str2double(code_curr);
                    else
                        cha = double(code_curr) - 55;
                        
                        if (cha > 9)
                            vals(vals_off) = floor(cha / 10);
                            vals_off = vals_off + 1;
                        end
                        
                        vals(vals_off) = cha;
                    end
                end
                
                vals = flipud(vals);
                
                vals_sum = 0;
                
                for j = 1:numel(vals)
                    if (mod((j - 1),2) == 0)
                        dou = vals(j) * 2;
                        vals_sum = vals_sum + floor(dou / 10) + mod(dou,10);
                    else
                        vals_sum = vals_sum + vals(j);
                    end
                end
                
                res(i) = mod((10 - mod(vals_sum,10)),10) == str2double(isin_curr(end));
            end
        end
    end
end