package baseltools;

import java.awt.Component;
import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererOpCapRslt extends RendererDefault
{
	private static final long serialVersionUID = 1L;

	private static final int COLUMN_HEADER = 0;
	private static final int ROW_ILM = 2;
	private static final int ROW_K_SMA = 3;
	private static final int ROW_RWA_SMA = 4;
	private static final int ROW_K_B2 = 5;
	private static final int ROW_RWA_B2 = 6;
	
	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		switch (row)
		{
			case ROW_K_SMA:
			case ROW_RWA_SMA:
				cell.setBackground(Environment.ColorOpCapSMA);
				break;
				
			case ROW_K_B2:
			case ROW_RWA_B2:
				cell.setBackground(Environment.ColorOpCapB2);
				break;

			default:
				cell.setBackground(Environment.ColorOpCapAll);
				break;
		}

		cell.setHorizontalAlignment(SwingConstants.CENTER);
		
		if (column == COLUMN_HEADER)
			cell.setFont(cell.getFont().deriveFont(20.0f));
		else if (value instanceof Double)
		{
			if (row == ROW_ILM)
				cell.setText(Environment.FormatNumber(value, true, 4));
			else
				cell.setText(Environment.FormatCurrencyGrouped.format(value));
		}

		return cell;
	}
}