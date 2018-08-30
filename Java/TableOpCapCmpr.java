package baseltools;

public final class TableOpCapCmpr extends TableDefault
{
	private static final long serialVersionUID = 1L;
	
	public TableOpCapCmpr(final Object[][] data, final Object[] columnNames, final boolean[] columnEditing)
	{
		super(data, columnNames, columnEditing);
	}

	@Override
	public boolean isCellEditable(final int row, final int column)
	{
		return (column > 0);
	}
}