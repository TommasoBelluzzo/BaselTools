package baseltools;

import java.awt.Color;
import java.awt.Component;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;
import javax.swing.UIManager;

import com.jidesoft.grid.CellSpan;

public final class RendererCcrRsltNets extends RendererDefault
{
	private static final long serialVersionUID = 1L;

	private static final int COLUMN_ID = 0;
	private static final int COLUMN_EAD_FINAL = 1;
	private static final int COLUMN_EAD = 7;
	
	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
		
		final boolean columnGlobal = (column <= COLUMN_EAD_FINAL);
		final boolean rowEven = ((row % 2) == 0);
		
		final TableDefault model = (TableDefault)table.getModel();
		
		final int selectedRow;
		final Object tag = model.getTag();
		
		if (tag == null)
			selectedRow = -99;
		else
			selectedRow = (Integer)tag;
		
		if (rowEven && columnGlobal && (model.getCellSpanAt(row, column) == null))
			model.addCellSpan(new CellSpan(row, column, 2, 1));

		if ((row == selectedRow) || (row == (selectedRow - 1)))
		{
			final Color colorSelection = UIManager.getColor("Table.selectionBackground");
			cell.setBackground(colorSelection);
		}
		else
		{
			final String id;

			if (rowEven)
			{
				if (column == COLUMN_ID)
					id = (String)value;
				else
					id = (String)table.getValueAt(row, COLUMN_ID);
			}
			else
				id = (String)table.getValueAt((row - 1), COLUMN_ID);
			
			final String ead;
			
			if (column == COLUMN_EAD)
				ead = (String)value;
			else
				ead = (String)table.getValueAt(row, COLUMN_EAD);
			
			if (columnGlobal || !ead.endsWith("N/A"))
			{
				if (id.endsWith("N"))
					cell.setBackground(Environment.ColorCcrNetsStd);
				else
					cell.setBackground(Environment.ColorCcrNetsTrd);
			}
			else
				cell.setBackground(Environment.ColorGray);
		}
		
		cell.setHorizontalAlignment(SwingConstants.CENTER);

		return cell;
	}
}