package baseltools;

import java.awt.Component;
import java.awt.Font;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererOpCapCmpr extends RendererDefault
{
	private static final long serialVersionUID = 1L;

	private static final int COLUMN_HEADER = 0;
	
	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		cell.setHorizontalAlignment(SwingConstants.CENTER);
		
		if (column == COLUMN_HEADER)
		{
			cell.setBackground(Environment.ColorOpCapB2);
			cell.setFont(new Font(cell.getFont().getName(), Font.BOLD, 17));
		}
		else
		{
			cell.setBackground(Environment.ColorEnabled);
			cell.setText(Environment.FormatCurrencyGrouped.format(value));
		}

		return cell;
	}
}