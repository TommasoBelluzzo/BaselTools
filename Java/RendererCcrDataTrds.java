package baseltools;

import java.awt.Component;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererCcrDataTrds extends RendererDefault
{
	private static final long serialVersionUID = 1L;
	
	private static final int COLUMN_ID = 0;
	private static final int COLUMN_CLASS = 2;
	private static final int COLUMN_NOTIONAL = 3;
	private static final int COLUMN_VALUE = 4;
	private static final int COLUMN_START = 5;
	private static final int COLUMN_END = 6;
	private static final int COLUMN_MATURITY = 7;
	private static final int COLUMN_OPTION = 10;

	private final transient String[] Formats;

	public RendererCcrDataTrds(final int[] maximums)
	{
		super();

		final int maximumsLength = maximums.length;
		final int lastColumn = COLUMN_OPTION + 1;
		
		if (maximumsLength != lastColumn)
			throw new IllegalArgumentException("The 'maximums' argument must contain a number of elements equal to the number of table columns.");

		Formats = new String[lastColumn];
		
		for (int i = 0; i < maximumsLength; ++i)
		{
			final int maximumCurrent = maximums[i];
			
			if (maximumCurrent == -1)
				Formats[i] = "%1$s";
			else
				Formats[i] = "%1$" + maximumCurrent + "s";
		}
	}
	
	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		String cls;
		
		if (column == COLUMN_CLASS)
			cls = (String)value;
		else
			cls = (String)table.getValueAt(row, COLUMN_CLASS);

		if (cls.startsWith("CO"))
			cell.setBackground(Environment.ColorCcrTrdsCo);
		else if (cls.startsWith("CR"))
			cell.setBackground(Environment.ColorCcrTrdsCr);
		else if (cls.startsWith("EQ"))
			cell.setBackground(Environment.ColorCcrTrdsEq);
		else if (cls.startsWith("FX"))
			cell.setBackground(Environment.ColorCcrTrdsFx);
		else if (cls.startsWith("IR"))
			cell.setBackground(Environment.ColorCcrTrdsIr);
		else
			cell.setBackground(Environment.ColorCcrTrdsRe);
		
		cell.setHorizontalAlignment(SwingConstants.CENTER);
		
		final String text;
		
		switch (column)
		{
			case COLUMN_ID:
				text = Integer.toString((Integer)value);
				break;

			case COLUMN_NOTIONAL:
			case COLUMN_VALUE:
				text = Environment.FormatCurrencyGrouped.format(value);
				break;
				
			case COLUMN_START:
			case COLUMN_END:
			case COLUMN_MATURITY:
				text = Environment.FormatNumber(value,false,2);
				break;
				
			case COLUMN_OPTION:
			{
				final boolean option = (Boolean)value;

				if (option)
					text = "YES";
				else
					text = "NO";
				
				break;
			}
				
			default:
				text = (String)value;
				break;
		}

		cell.setText(String.format(Formats[column], text));

		return cell;
	}
}