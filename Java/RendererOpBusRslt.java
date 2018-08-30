package baseltools;

import java.awt.Component;
import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererOpBusRslt extends RendererDefault
{
	private static final long serialVersionUID = 1L;
	
	private static final int COLUMN_HEADER = 0;
	private static final int ROW_ILDC = 0;
	private static final int ROW_FC = 1;
	private static final int ROW_SC = 2;
	private static final int ROW_UBI = 3;

	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		switch (row)
		{
			case ROW_ILDC:
				cell.setBackground(Environment.ColorOpBusILDC);
				break;
				
			case ROW_FC:
				cell.setBackground(Environment.ColorOpBusFC);
				break;
				
			case ROW_SC:
				cell.setBackground(Environment.ColorOpBusSC);
				break;
				
			case ROW_UBI:
				cell.setBackground(Environment.ColorOpBusUBI);
				break;
				
			default:
				cell.setBackground(Environment.ColorOpBusBI);
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