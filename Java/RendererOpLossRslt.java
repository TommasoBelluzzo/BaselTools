package baseltools;

import java.awt.Component;
import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererOpLossRslt extends RendererDefault
{
	private static final long serialVersionUID = 1L;
	
	private static final int COLUMN_HEADER = 0;
	private static final int ROW_LOSSALL = 0;
	private static final int ROW_LOSS10 = 1;
	private static final int ROW_LOSS100 = 2;

	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		switch (row)
		{
			case ROW_LOSSALL:
				cell.setBackground(Environment.ColorOpLossAll);
				break;
				
			case ROW_LOSS10:
				cell.setBackground(Environment.ColorOpLoss10);
				break;
				
			case ROW_LOSS100:
				cell.setBackground(Environment.ColorOpLoss100);
				break;
				
			default:
				cell.setBackground(Environment.ColorOpLossLC);
				break;
		}
		
		cell.setHorizontalAlignment(SwingConstants.CENTER);
		
		if (column == COLUMN_HEADER)
			cell.setFont(cell.getFont().deriveFont(20.0f));
		else
			cell.setText(Environment.FormatCurrencyGrouped.format(value));

		return cell;
	}
}