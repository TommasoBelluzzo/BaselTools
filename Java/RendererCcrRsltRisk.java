package baseltools;

import java.awt.Component;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

import com.jidesoft.grid.CellSpan;

public final class RendererCcrRsltRisk extends RendererDefault
{
	private static final long serialVersionUID = 1L;

	private static final int COLUMN_NAME = 0;
	private static final int COLUMN_EFFNOT = 2;
	private static final int COLUMN_ADDON = 3;
	
	private final transient String HedgingSet;

	public RendererCcrRsltRisk(final String hedgingSet)
	{
		super();

		HedgingSet = hedgingSet;
	}
	
	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		final boolean columnGlobal = (column == COLUMN_NAME);
		final boolean rowEven = ((row % 2) == 0);
		
		if (rowEven && columnGlobal)
		{
			final TableDefault model = (TableDefault)table.getModel();
			
			if (model.getCellSpanAt(row, column) == null)
				model.addCellSpan(new CellSpan(row, column, 2, 1));
			
			cell.setFont(cell.getFont().deriveFont(14.0f));
		}
		
		if (columnGlobal)
		{
			if (HedgingSet.startsWith("COMMODITIES"))
				cell.setBackground(Environment.ColorCcrTrdsCo);
			else if (HedgingSet.startsWith("CREDIT"))
				cell.setBackground(Environment.ColorCcrTrdsCr);
			else if (HedgingSet.startsWith("EQUITY"))
				cell.setBackground(Environment.ColorCcrTrdsEq);
			else if (HedgingSet.startsWith("FOREIGN EXCHANGE"))
				cell.setBackground(Environment.ColorCcrTrdsFx);
			else if (HedgingSet.startsWith("INTEREST RATE"))
				cell.setBackground(Environment.ColorCcrTrdsIr);
			else
				cell.setBackground(Environment.ColorCcrTrdsRe);
		}
		else
		{
			if (HedgingSet.startsWith("INTEREST RATE"))
			{
				if (columnGlobal)
					cell.setBackground(Environment.ColorCcrTrdsIr);
				else if (column <= COLUMN_ADDON)
				{
					final String en = (String)table.getValueAt(row, COLUMN_EFFNOT);
					
					if (en.endsWith("N/A"))
						cell.setBackground(Environment.ColorGray);
					else
						cell.setBackground(Environment.ColorCcrTrdsIr);
				}
				else
					cell.setBackground(Environment.ColorGray);
			}
			else
			{
				final String en = (String)table.getValueAt(row, COLUMN_EFFNOT);
				
				if (en.endsWith("N/A"))
					cell.setBackground(Environment.ColorGray);
				else
				{
					if (HedgingSet.startsWith("COMMODITIES"))
						cell.setBackground(Environment.ColorCcrTrdsCo);
					else if (HedgingSet.startsWith("CREDIT"))
						cell.setBackground(Environment.ColorCcrTrdsCr);
					else if (HedgingSet.startsWith("EQUITY"))
						cell.setBackground(Environment.ColorCcrTrdsEq);
					else if (HedgingSet.startsWith("FOREIGN EXCHANGE"))
						cell.setBackground(Environment.ColorCcrTrdsFx);
					else if (HedgingSet.startsWith("INTEREST RATE"))
						cell.setBackground(Environment.ColorCcrTrdsIr);
					else
						cell.setBackground(Environment.ColorCcrTrdsRe);
				}
			}
		}

		cell.setHorizontalAlignment(SwingConstants.CENTER);

		return cell;
	}
}