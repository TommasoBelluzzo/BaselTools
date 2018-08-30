package baseltools;

public final class TableOpBusData extends TableDefault
{
	private static final long serialVersionUID = 1L;
	
	public TableOpBusData(final Object[][] data, final Object[] columnNames, final boolean[] columnEditing)
	{
		super(data, columnNames, columnEditing);
	}

	@Override
	public boolean isCellEditable(final int row, final int column)
	{
		return ((column > 0) && (column < 4) && (row != 0) && (row != 7) && (row != 10));
	}
}