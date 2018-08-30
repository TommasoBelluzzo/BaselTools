package baseltools;

import java.util.Arrays;

import com.jidesoft.grid.DefaultSpanTableModel;

public class TableDefault extends DefaultSpanTableModel
{
	private static final long serialVersionUID = 1L;
	
	private final transient boolean[] ColumnEditing;
    private Object Tag;

	public TableDefault(final Object[][] data, final Object[] columnNames, final boolean[] columnEditing)
	{
		super(data, columnNames);

		ColumnEditing = Arrays.copyOf(columnEditing, columnEditing.length);
	}

	@Override
	public boolean isCellEditable(final int row, final int column)
	{
		return ColumnEditing[column];
	}
	
    public Object getTag()
    {
    	return Tag;
    }

	public Object[][] getData()
	{
		final int cols = getColumnCount();
		final int rows = getRowCount();

		Object[][] data = new Object[rows][cols];

	    for (int i = 0; i < rows; ++i)
	    {
	        for (int j = 0; j < cols; ++j)
	        	data[i][j] = getValueAt(i, j);
	    }

	    return data;
	}
	
	public void setTag(final Object value)
	{
		Tag = value;
	}
}