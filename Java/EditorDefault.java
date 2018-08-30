package baseltools;

import javax.swing.DefaultCellEditor;
import javax.swing.JFormattedTextField;

public final class EditorDefault extends DefaultCellEditor
{
	private static final long serialVersionUID = 1L;

	public EditorDefault()
	{
		super(new JFormattedTextField());
	}
}