classdef (Abstract) BaselInterface < handle
    %% Methods: Events
    methods (Access = protected)
        function TableColumnHeader_Clicked(this,obj,evd) %#ok<INUSL>
            import('javax.swing.SwingUtilities');
            
            sort = obj.isSortable();
            
            if (~sort && ~SwingUtilities.isLeftMouseButton(evd))
                return;
            elseif (sort && ~SwingUtilities.isRightMouseButton(evd))
                return;
            end
            
            col = obj.columnAtPoint(evd.getPoint());
            col_idx = obj.convertColumnIndexToModel(col);
            
            sel_mod = obj.getTableSelectionModel();
            
            if (evd.isControlDown())
                for i = 0:obj.getRowCount()-1
                    sel_mod.addSelection(i,col_idx);
                    
                    drawnow();
                    pause(0.01);
                end
            else
                sel_mod.setSelectionInterval(0,(obj.getRowCount() - 1),col_idx);
            end
        end
        
        function TableRowHeader_Clicked(this,obj,evd) %#ok<INUSL>
            import('javax.swing.SwingUtilities');
            
            if (~SwingUtilities.isLeftMouseButton(evd))
                return;
            end
            
            row = obj.rowAtPoint(evd.getPoint());
            row_idx = obj.convertRowIndexToModel(row);
            
            if (~evd.isControlDown())
                obj.clearSelection();
                
                drawnow();
                pause(0.01);
            end
            
            sel_mod = obj.getTableSelectionModel();
            
            for i = 0:obj.getColumnCount()-1
                sel_mod.addSelection(row_idx,i);
                
                drawnow();
                pause(0.01);
            end
        end
    end
    
    %% Methods: Functions
    methods (Access = protected)
        function res = FormatException(this,msg,e) %#ok<INUSL>
            e_msg = strtrim(e.message);
            
            if (startsWith(e_msg,'"'))
                e_msg = e_msg(2:end);
            end
            
            if (endsWith(e_msg,'"'))
                e_msg = e_msg(1:end-1);
            end
            
            e_msg = regexprep(e_msg,'[\n\r]+',' ');
            e_msg = strrep(e_msg,'"','''');

            e_ref = {};
            
            for i = 1:numel(e.stack)
                e_stk = e.stack(i);
                e_stk_name = e_stk.name;
                
                if (startsWith(e_stk_name,'Basel'))
                    e_fun_spl = strsplit(e_stk_name,'.');
                    e_ref = {char(e_fun_spl(end)) num2str(e_stk.line)};
                    
                    break;
                end
            end
            
            if (isempty(e_ref))
                e_stk = e.stack(1);
                e_stk_name = e_stk.name;

                e_fun_spl = strsplit(e_stk_name,'.');
                e_ref = {char(e_fun_spl(end)) num2str(e_stk.line)};
            end
            
            if (isempty(msg))
                res = ['The exception produced is: "' e_msg '"' ' (' e_ref{1} ', ' e_ref{2} ').'];
            else
                res = [msg ' ' 'The exception produced is: "' e_msg '"' ' (' e_ref{1} ', ' e_ref{2} ').'];
            end
        end
        
        function SetupBox(this,box,varargin) %#ok<INUSL>
            persistent ip;
            
            if (isempty(ip))
                ip = inputParser();
                ip.CaseSensitive = true;
                ip.addParameter('VerticalScrollbar',false,@(x)validateattributes(x,{'logical'},{'scalar'}));
            end
            
            if (isempty(box.UserData))
                return;
            end
            
            ip.parse(varargin{:});
            ip_res = ip.Results;
            
            html = box.UserData{1};
            box.UserData = [];
            
            import('javax.swing.*');
            import('javax.swing.text.html.*');
            import('baseltools.*');
            
            jpan = javaObjectEDT(findjobj(box));
            jbox = jpan.getViewport().getView();
            
            jpan.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            
            if (ip_res.VerticalScrollbar)
                jpan.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
            else
                jpan.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
            end
            
            if (strcmp(box.Enable,'inactive'))
                jpan.setBorder(BorderFactory.createEmptyBorder());
                
                box_x = box.Position(1);
                box_y = box.Position(2);
                box_hei = box.Position(4);
                
                pan1 = uipanel(box.Parent);
                pan1.BorderType = 'none';
                pan1.BorderWidth = 0;
                pan1.Units = 'pixels';
                pan1.Position = [box_x box_y 2 box_hei];
                
                pan2 = uipanel(box.Parent);
                pan2.BorderType = 'none';
                pan2.BorderWidth = 0;
                pan2.Units = 'pixels';
                pan2.Position = [(box_x + box.Position(3) - 2) box_y 2 box_hei];
            end
            
            jbox.setDoubleBuffered(true);
            jbox.setEditable(false);
            jbox.setEditorKit(HTMLEditorKit());
            jbox.addHyperlinkListener(Environment.HyperlinkBrowser);
            jbox.setText(html);
        end
        
        function SetupTable(this,tab,varargin)
            persistent ip;
            
            if (isempty(ip))
                ip = inputParser();
                ip.CaseSensitive = true;
                
                ip.addParameter('Data',[],@(x)validateattributes(x,{'cell'},{'2d','nonempty'}));
                ip.addParameter('DataChanged',[],@(x)validateattributes(x,{'function_handle'},{'scalar'}));
                ip.addParameter('Editor',{'EditorDefault'},@(x)validateattributes(x,{'cell'},{'vector'}));
                ip.addParameter('FillHeight',true,@(x)validateattributes(x,{'logical'},{'scalar'}));
                ip.addParameter('MouseClicked',[],@(x)validateattributes(x,{'function_handle'},{'scalar'}));
                ip.addParameter('MouseEntered',[],@(x)validateattributes(x,{'function_handle'},{'scalar'}));
                ip.addParameter('MouseExited',[],@(x)validateattributes(x,{'function_handle'},{'scalar'}));
                ip.addParameter('Renderer',{'RendererDefault'},@(x)validateattributes(x,{'cell'},{'vector'}));
                ip.addParameter('RowHeaderWidth',0,@(x)validateattributes(x,{'numeric'},{'scalar','integer','real','finite','>=',1}));
                ip.addParameter('RowsHeight',0,@(x)validateattributes(x,{'numeric'},{'scalar','integer','real','finite','>=',1}));
                ip.addParameter('Selection',[false false false],@(x)validateattributes(x,{'logical'},{'vector','numel',3}));
                ip.addParameter('Sorting',0,@(x)validateattributes(x,{'numeric'},{'scalar','integer','real','finite','>=',0,'<=',2}));
                ip.addParameter('Table',{'TableDefault'},@(x)validateattributes(x,{'cell'},{'vector'}));
                ip.addParameter('VerticalScrollbar',false,@(x)validateattributes(x,{'logical'},{'scalar'}));
            end
            
            ip.parse(varargin{:});
            ip_res = ip.Results;
            
            tab.ColumnFormat = [];
            
            import('java.awt.*');
            import('javax.lang.*');
            import('javax.swing.*');
            import('com.jidesoft.grid.*');
            import('baseltools.*');
            
            jpan = javaObjectEDT(findjobj(tab,'Persist'));
            jvp = javaObjectEDT(jpan.getViewport());
            jtab = javaObjectEDT(jvp.getView());
            col_mod = javaObjectEDT(jtab.getColumnModel());
            hea_col = javaObjectEDT(jtab.getTableHeader());
            
            if (isempty(tab.RowName))
                hea_row = javaObjectEDT(JViewport());
            else
                hea_row = javaObjectEDT(jpan.getRowHeader());
            end
            
            if (numel(ip_res.Editor) == 1)
                tab_edi = javaObjectEDT(eval([ip_res.Editor{1} '()']));
            else
                tab_edi = javaObjectEDT(eval([ip_res.Editor{1} '(ip_res.Editor{2:end})']));
            end
            
            if (numel(ip_res.Table) == 1)
                tab_mod = eval([ip_res.Table{1} '(ip_res.Data,tab.ColumnName,tab.ColumnEditable)']);
            else
                tab_mod = eval([ip_res.Table{1} '(ip_res.Data,tab.ColumnName,tab.ColumnEditable,ip_res.Table{2:end})']);
            end
            
            if (numel(ip_res.Renderer) == 1)
                tab_ren = javaObjectEDT(eval([ip_res.Renderer{1} '()']));
            else
                tab_ren = javaObjectEDT(eval([ip_res.Renderer{1} '(ip_res.Renderer{2:end})']));
            end
            
            jpan.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
            
            if (ip_res.VerticalScrollbar)
                jpan.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
            else
                jpan.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
            end
            
            jtab.putClientProperty('terminateEditOnFocusLost',true);
            jtab.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
            jtab.setDoubleBuffered(true);
            jtab.setFillsViewportHeight(true);
            jtab.setKeepColumnAtPoint(true);
            jtab.setKeepRowAtPoint(true);
            jtab.setTableColumnWidthKeeper(DefaultTableColumnWidthKeeper());
            
            if (isempty(ip_res.Data))
                ip_res.Data = cell(0,numel(tab.ColumnName));
                
                clr = Environment.ColorDisabled;
                tab.Enable = 'off';
            else
                clr = Environment.ColorEnabled;
                tab.Enable = 'on';
            end
            
            jvp.setBackground(clr);
            hea_col.setBackground(clr);
            hea_row.setBackground(clr);
            
            jtab.setModel(tab_mod);
            
            if (~isempty(ip_res.DataChanged))
                cp = handle(tab_mod,'CallbackProperties');
                set(cp,'TableChangedCallback',@(obj,evd)ip_res.DataChanged(jtab,evd));
            end
            
            mou_cli = ~isempty(ip_res.MouseClicked);
            mou_ent = ~isempty(ip_res.MouseEntered);
            mou_exi = ~isempty(ip_res.MouseExited);
            
            if (mou_cli || mou_ent || mou_exi)
                cp = handle(jtab,'CallbackProperties');
                
                if (mou_cli)
                    set(cp,'MouseClickedCallback',@(obj,evd)ip_res.MouseClicked(jtab,evd));
                end
                
                if (mou_ent)
                    set(cp,'MouseEnteredCallback',@(obj,evd)ip_res.MouseEntered(jtab,evd));
                end
                
                if (mou_exi)
                    set(cp,'MouseExitedCallback',@(obj,evd)ip_res.MouseExited(jtab,evd));
                end
            end
            
            if (strcmp(tab.Enable,'off'))
                hea_col.setEnabled(false);
            else
                hea_col.setEnabled(true);
            end
            
            if (ip_res.Sorting == 0)
                jtab.setSortable(false);
            else
                jtab.setSortable(true);
                jtab.setAutoResort(true);
                jtab.setPreserveSelectionsAfterSorting(true);
                
                if (ip_res.Sorting == 1)
                    jtab.setMultiColumnSortable(false);
                else
                    jtab.setMultiColumnSortable(true);
                end
            end
            
            jtab.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
            jtab.setCellSelectionEnabled(ip_res.Selection(1));
            jtab.setNonContiguousCellSelection(ip_res.Selection(1));
            jtab.setColumnSelectionAllowed(ip_res.Selection(2));
            jtab.setRowSelectionAllowed(ip_res.Selection(3));
            
            if (ip_res.Selection(2))
                if (ip_res.Sorting == 0)
                    hea_col.setToolTipText('<html><body><b>LClick:</b> Column Selection<br><b>CTRL+LClick:</b> Add Column to Current Selection</body></html>');
                else
                    hea_col.setToolTipText('<html><body><b>(CTRL+)LClick1:</b> Sort (Secondary) Ascending<br><b>(CTRL+)LClick2:</b> Sort (Secondary) Descending<br><b>(CTRL+)LClick3:</b> Unsort (Secondary)<br><br><b>RClick:</b> Column Selection<br><b>CTRL+RClick:</b> Add Column to Current Selection</body></html>');
                end
                
                cp = handle(hea_col,'CallbackProperties');
                set(cp,'MousePressedCallback',@(obj,evd)this.TableColumnHeader_Clicked(jtab,evd));
            end
            
            if (~isempty(tab.RowName))
                hea_row_view = javaObjectEDT(hea_row.getView());
                
                if (ip_res.RowHeaderWidth > 0)
                    hea_row_siz = Dimension(ip_res.RowHeaderWidth,hea_row_view.getHeight());
                    
                    hea_row.setPreferredSize(hea_row_siz);
                    hea_row.setSize(hea_row_siz);
                end
                
                if (ip_res.Selection(3))
                    hea_row_view.setToolTipText('<html><body><b>LClick:</b> Row Selection<br><b>CTRL+LClick:</b> Add Row to Current Selection</body></html>');
                    
                    cp = handle(hea_row_view,'CallbackProperties');
                    set(cp,'MousePressedCallback',@(obj,evd)this.TableRowHeader_Clicked(jtab,evd));
                end
            end
            
            if (ip_res.RowsHeight > 0)
                jtab.setRowHeight(ip_res.RowsHeight);
                
                if (ip_res.FillHeight)
                    rows = size(ip_res.Data,1);
                    
                    if (rows > 0)
                        hei_diff = tab.Position(4) - (rows * ip_res.RowsHeight) - 29;
                        
                        if (hei_diff > 0)
                            if (hei_diff < rows)
                                rows_off = 0;
                                
                                while (hei_diff > 0)
                                    jtab.setRowHeight(rows_off,(ip_res.RowsHeight + 1));
                                    rows_off = rows_off + 1;
                                    
                                    hei_diff = hei_diff - 1;
                                end
                            else
                                jtab.setRowHeight((rows - 1),(ip_res.RowsHeight + hei_diff));
                            end
                        end
                    end
                end
            end
            
            for i = 1:jtab.getColumnCount()
                col = javaObjectEDT(col_mod.getColumn(i - 1));
                col.setCellEditor(tab_edi);
                col.setCellRenderer(tab_ren);
                
                col_wid = tab.ColumnWidth{i};
                
                if (~ischar(col_wid))
                    col.setWidth(col_wid);
                end
                
                col.setResizable(false);
            end
            
            jtab.repaint();
        end
        
        function SetupTextbox(this,tb) %#ok<INUSL>
            import('java.awt.*');
            
            jtb = javaObjectEDT(findjobj(tb));
            jtb.setEditable(false);
            jtb.getCaret().setSelectionVisible(true);
        end
    end
end