package baseltools;

import java.awt.Component;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererCcrDataCols extends RendererDefault
{
	private static final long serialVersionUID = 1L;
	
	private static final int COLUMN_ID = 0;
	private static final int COLUMN_MARGIN = 3;
	private static final int COLUMN_VALUE = 4;
	private static final int COLUMN_PARAMETER = 5;
	private static final int COLUMN_MATURITY = 6;

	private final transient String[] Formats;

	public RendererCcrDataCols(final int[] maximums)
	{
		super();

		final int maximumsLength = maximums.length;
		final int lastColumn = COLUMN_MATURITY + 1;
		
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

		final String margin;
		
		if (column == COLUMN_MARGIN)
			margin = (String)value;
		else
			margin = (String)table.getValueAt(row, COLUMN_MARGIN);

		if ("ICA".equals(margin))
			cell.setBackground(Environment.ColorCcrColsICA);
		else
			cell.setBackground(Environment.ColorCcrColsVM);
		
		cell.setHorizontalAlignment(SwingConstants.CENTER);

		final String text;
		
		switch (column)
		{
			case COLUMN_ID:
				text = Integer.toString((Integer)value);
				break;
				
			case COLUMN_VALUE:
				text = Environment.FormatCurrencyGrouped.format(value);
				break;
				
			case COLUMN_PARAMETER:
			{
				final String valueText = (String)value;
				
				if ((valueText == null) || valueText.equals(""))
					text = "N/A";
				else
					text = valueText;
				
				break;
			}
				
			case COLUMN_MATURITY:
			{
				final Double number = (Double)value;
				
				if (number.isNaN())
					text = "N/A";
				else
					text = Environment.FormatNumber(value,false,2);
				
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