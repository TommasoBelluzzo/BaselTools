package baseltools;

import java.awt.event.ActionEvent;
import javax.swing.AbstractAction;

public final class EditorCurrencyAction extends AbstractAction
{
	private static final long serialVersionUID = 1L;

	private final transient EditorCurrency Editor;

	public EditorCurrencyAction(final EditorCurrency editor)
	{
		super();
		
		Editor = editor;
	}
	
	@Override
	public void actionPerformed(final ActionEvent evt)
	{
		if (Editor.validEdit())
		{
			try
			{
				Editor.signalEdit(true);
			}
			catch (Exception e) { }
		}
		else if (Editor.revertEdit())
		{
			try
			{
				Editor.signalEdit(false);
			}
			catch (Exception e) {  }
		}
	}
}
