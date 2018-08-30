package baseltools;

import java.awt.Component;
import java.util.Date;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererOpLossData extends RendererDefault
{
	private static final long serialVersionUID = 1L;

	private static final int COLUMN_ID = 0;
	private static final int COLUMN_DATE = 1;
	private static final int COLUMN_LOSS = 4;
	private static final int COLUMN_RECOVERY = 5;
	private static final int VALUE_LOSS10 = 10000000;
	private static final int VALUE_LOSS100 = 100000000;

	private final transient String[] Formats;
	
	public RendererOpLossData(final int[] maximums)
	{
		super();

		final int maximumsLength = maximums.length;
		final int lastColumn = COLUMN_RECOVERY + 1;
		
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
		
		final Double loss = (Double)table.getValueAt(row, COLUMN_LOSS);

		if (loss <= VALUE_LOSS10)
			cell.setBackground(Environment.ColorOpLossAll);
		else if (loss < VALUE_LOSS100)
			cell.setBackground(Environment.ColorOpLoss10);
		else
			cell.setBackground(Environment.ColorOpLoss100);

		cell.setHorizontalAlignment(SwingConstants.CENTER);

		final String text;
		
		switch (column)
		{
			case COLUMN_ID:
				text = Integer.toString((Integer)value);
				break;
				
			case COLUMN_DATE:
			{
				final Double number = (Double)value;
				final long millis = (long)((number - 719529.0d) * 86400000.0d);
				final Date date = new Date(millis);
				
				text = Environment.DateFormat.format(date);

				break;
			}

			case COLUMN_LOSS:
			case COLUMN_RECOVERY:
				text = Environment.FormatCurrencyGrouped.format(value);
				break;

			default:
				text = (String)value;
				break;
		}
		
		cell.setText(String.format(Formats[column], text));

		return cell;
	}
}