package baseltools;

import java.awt.Component;
import java.awt.Font;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererOpBusComp extends RendererDefault
{
	private static final long serialVersionUID = 1L;

	private static final int COLUMN_VALUE = 2;
	
	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		cell.setHorizontalAlignment(SwingConstants.CENTER);
		
		if (column == COLUMN_VALUE)
		{			
			if (value instanceof Double)
	    	{
				cell.setBackground(Environment.ColorOpBusBIC);
	    		cell.setFont(cell.getFont().deriveFont(Font.BOLD));
	    		cell.setText(Environment.FormatCurrencyGrouped.format(value));
	    	}
			else
			{
				cell.setBackground(Environment.ColorDisabled);
				cell.setFont(cell.getFont().deriveFont(Font.PLAIN));
			}
		}
		else
		{
			final Font font = cell.getFont();
			final Object bic = table.getValueAt(row, COLUMN_VALUE);

			int fontStyle;
			
			if (bic instanceof Double)
	    	{
	    		fontStyle = Font.BOLD;
				cell.setBackground(Environment.ColorOpBusBIC);
	    	}
			else
			{
				fontStyle = Font.PLAIN;
				cell.setBackground(Environment.ColorDisabled);
			}
			
			if (column == 0)
				cell.setFont(new Font(font.getName(), fontStyle, 20));
			else
				cell.setFont(new Font(font.getName(), fontStyle, font.getSize()));
		}

		return cell;
	}
}