package baseltools;

import java.awt.Component;
import java.awt.Font;
import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.SwingConstants;

import com.jidesoft.grid.CellSpan;

public final class RendererOpBusData extends RendererDefault
{
	private static final long serialVersionUID = 1L;

	private static final int COLUMN_HEADER = 0;
	private static final int COLUMN_AVERAGE = 4;
	private static final int ROW_ILDC = 0;
	private static final int ROW_FC = 7;
	private static final int ROW_SC = 10;
	private static final int ROWS_HEIGHT = 28;

	@Override
	public Component getTableCellRendererComponent(final JTable table, final Object value, final boolean isSelected, final boolean hasFocus, final int row, final int column)
	{
		final JLabel cell = (JLabel)super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

		cell.setHorizontalAlignment(SwingConstants.CENTER);
		
		if (column == COLUMN_HEADER)
		{
			if (row < ROW_FC)
				cell.setBackground(Environment.ColorOpBusILDC);
			else if (row < ROW_SC)
				cell.setBackground(Environment.ColorOpBusFC);
			else
				cell.setBackground(Environment.ColorOpBusSC);

			if ((row == ROW_ILDC) || (row == ROW_FC) || (row == ROW_SC))
			{
				final Font font = cell.getFont();
				final TableOpBusData model = (TableOpBusData)table.getModel();
				
				if (table.getRowHeight(row) != ROWS_HEIGHT)
					table.setRowHeight(row, ROWS_HEIGHT);

				if (model.getCellSpanAt(row, COLUMN_HEADER) == null)
					model.addCellSpan(new CellSpan(row, COLUMN_HEADER, 1, 5));

				cell.setBorder(Environment.BorderHeaderColumn);
				cell.setFont(new Font(font.getName(), Font.BOLD, 16));
			}
			else
			{
				cell.setBorder(Environment.BorderHeaderRow);
				cell.setFont(cell.getFont().deriveFont(Font.BOLD));
			}
		}
		else
		{
			if (column == COLUMN_AVERAGE)
			{
				cell.setBackground(Environment.ColorDisabled);
				cell.setFont(cell.getFont().deriveFont(Font.BOLD));
			}
			else
				cell.setBackground(Environment.ColorEnabled);
		}

		if (value instanceof Double)
    		cell.setText(Environment.FormatCurrencyGrouped.format(value));
		
		return cell;
	}
}