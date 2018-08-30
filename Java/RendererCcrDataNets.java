package baseltools;

import java.awt.Component;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

public final class RendererCcrDataNets extends RendererDefault
{
	private static final long serialVersionUID = 1L;
	
	private static final int COLUMN_ID = 0;
	private static final int COLUMN_MARGINED = 1;
	private static final int COLUMN_THRESHOLD = 2;
	private static final int COLUMN_MTA = 3;
	private static final int COLUMN_MPOR = 4;
	
	private final transient String[] Formats;

	public RendererCcrDataNets(final int[] maximums)
	{
		super();

		final int maximumsLength = maximums.length;
		final int lastColumn = COLUMN_MPOR + 1;
		
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

		String id;
		
		if (column == COLUMN_ID)
			id = (String)value;
		else
			id = (String)table.getValueAt(row, COLUMN_ID);

		if (id.endsWith("N"))
			cell.setBackground(Environment.ColorCcrNetsStd);
		else
			cell.setBackground(Environment.ColorCcrNetsTrd);
		
		cell.setHorizontalAlignment(SwingConstants.CENTER);

		final String text;
		
		switch (column)
		{
			case COLUMN_MARGINED:
			{
				final boolean margined = (Boolean)value;

				if (margined)
					text = "YES";
				else
					text = "NO";
				
				break;
			}
			
			case COLUMN_THRESHOLD:
			case COLUMN_MTA:
			case COLUMN_MPOR:
			{
				final boolean margined = (Boolean)table.getValueAt(row, COLUMN_MARGINED);

				if (margined)
					text = Environment.FormatCurrencyGrouped.format(value);
				else
					text = "N/A";
				
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