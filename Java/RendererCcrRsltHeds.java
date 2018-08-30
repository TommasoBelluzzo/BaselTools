package baseltools;

import java.awt.Color;
import java.awt.Component;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;
import javax.swing.UIManager;

import com.jidesoft.grid.CellSpan;

public final class RendererCcrRsltHeds extends RendererDefault
{
	private static final long serialVersionUID = 1L;
	
	private static final int COLUMN_CLASS = 0;
	private static final int COLUMN_ADDON = 3;
	
	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		final boolean columnGlobal = (column == COLUMN_CLASS);
		final boolean rowEven = ((row % 2) == 0);
		
		final TableDefault model = (TableDefault)table.getModel();
		
		final int selectedRow;
		final Object tag = model.getTag();
		
		if (tag == null)
			selectedRow = -99;
		else
			selectedRow = (Integer)tag;
		
		if (rowEven && columnGlobal)
		{
			if (model.getCellSpanAt(row, column) == null)
				model.addCellSpan(new CellSpan(row, column, 2, 1));

			cell.setFont(cell.getFont().deriveFont(14.0f));
		}

		if ((row == selectedRow) || (row == (selectedRow - 1)))
		{
			final Color colorSelection = UIManager.getColor("Table.selectionBackground");
			cell.setBackground(colorSelection);
		}
		else
		{
			final String cls;
			
			if (rowEven)
			{
				if (column == COLUMN_CLASS)
					cls = (String)value;
				else
					cls = (String)table.getValueAt(row, COLUMN_CLASS);
			}
			else
				cls = (String)table.getValueAt((row - 1), COLUMN_CLASS);
			
			final String addon;
			
			if (column == COLUMN_ADDON)
				addon = (String)value;
			else
				addon = (String)table.getValueAt(row, COLUMN_ADDON);
		
			if (columnGlobal || !addon.endsWith("N/A"))
			{
				if (cls.startsWith("COMMODITIES"))
					cell.setBackground(Environment.ColorCcrTrdsCo);
				else if (cls.startsWith("CREDIT"))
					cell.setBackground(Environment.ColorCcrTrdsCr);
				else if (cls.startsWith("EQUITY"))
					cell.setBackground(Environment.ColorCcrTrdsEq);
				else if (cls.startsWith("FOREIGN EXCHANGE"))
					cell.setBackground(Environment.ColorCcrTrdsFx);
				else if (cls.startsWith("INTEREST RATE"))
					cell.setBackground(Environment.ColorCcrTrdsIr);
				else
					cell.setBackground(Environment.ColorCcrTrdsRe);
			}
			else
				cell.setBackground(Environment.ColorGray);
		}
		
		cell.setHorizontalAlignment(SwingConstants.CENTER);

		return cell;
	}
}