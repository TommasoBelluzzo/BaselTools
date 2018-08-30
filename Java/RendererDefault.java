package baseltools;

import javax.swing.SwingConstants;
import javax.swing.table.DefaultTableCellRenderer;

public class RendererDefault extends DefaultTableCellRenderer
{
	private static final long serialVersionUID = 1L;

	public RendererDefault()
	{
		super();

		setBackground(Environment.ColorEnabled);
		setForeground(Environment.ColorFont);
		setHorizontalAlignment(SwingConstants.CENTER);
	}
}