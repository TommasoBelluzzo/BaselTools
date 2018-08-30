package baseltools;

import java.awt.Component;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.ParsePosition;

import javax.swing.DefaultCellEditor;
import javax.swing.JFormattedTextField;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.KeyStroke;
import javax.swing.SwingUtilities;
import javax.swing.text.DefaultFormatterFactory;
import javax.swing.text.NumberFormatter;

public final class EditorCurrency extends DefaultCellEditor
{
	private static final long serialVersionUID = 1L;

	private static final int REVERT = 1;
	
	private final transient DecimalFormat FormatCurrencyUngrouped;
	private final transient double MaximumValue;
	private final transient double MinimumValue;
	private final transient JFormattedTextField TextField;
	private final transient String ErrorMessage;
	
	public EditorCurrency(final double minimumValue, final double maximumValue, final double currentValue)
	{
		super(new JFormattedTextField());

		FormatCurrencyUngrouped = (DecimalFormat)Environment.FormatCurrencyUngrouped.clone();
		MaximumValue = maximumValue;
		MinimumValue = minimumValue;
		
		final NumberFormatter formatter = new NumberFormatter();
        formatter.setFormat(FormatCurrencyUngrouped);
        formatter.setMaximum(maximumValue);
        formatter.setMinimum(minimumValue);

        TextField = (JFormattedTextField)getComponent();
        TextField.setFocusLostBehavior(JFormattedTextField.COMMIT_OR_REVERT);
        TextField.setFormatterFactory(new DefaultFormatterFactory(formatter));
        TextField.setValue(currentValue);
        TextField.getInputMap().put(KeyStroke.getKeyStroke(KeyEvent.VK_ENTER, 0), "Check");
        TextField.getActionMap().put("Check", (new EditorCurrencyAction(this)));
        
        final DecimalFormat formatCurrencyGrouped = (DecimalFormat)Environment.FormatCurrencyGrouped.clone();
        
        ErrorMessage = "The value must be a float between "
        		+ formatCurrencyGrouped.format(minimumValue) + " and "
                + formatCurrencyGrouped.format(maximumValue) + ".\n"
                + "You can either continue editing or revert to the last valid value.";
	}

    @Override
    public boolean stopCellEditing()
    {
        final JFormattedTextField editor = (JFormattedTextField)getComponent();

        if (validEdit())
        {
            try
            {
            	editor.commitEdit();
            }
            catch (Exception e) { }
        }
        else if (!revertEdit())
        	return false;

        return super.stopCellEditing();
    }

    @Override
    public Component getTableCellEditorComponent(final JTable table, final Object value, final boolean isSelected, final int row, final int column)
    {
    	final JFormattedTextField editor = (JFormattedTextField)super.getTableCellEditorComponent(table, value, isSelected, row, column);
    	final JLabel cell = (JLabel)table.getCellRenderer(row, column).getTableCellRendererComponent(table, value, isSelected, false, row, column);

    	editor.setFont(cell.getFont());
    	editor.setHorizontalAlignment(cell.getHorizontalAlignment());
    	editor.setText(FormatCurrencyUngrouped.format(value));
		editor.setValue(value);

        return editor;
    }

    @Override
    public Object getCellEditorValue()
    {
    	final JFormattedTextField editor = (JFormattedTextField)getComponent();
        final Object value = editor.getValue();

        if (value instanceof Double)
        	return value;

        if (value instanceof Number)
        	return (new Double(((Number)value).doubleValue()));

        return Double.NaN;
    }

	public boolean revertEdit()
	{
        Toolkit.getDefaultToolkit().beep();

        TextField.selectAll();

        final int answer = JOptionPane.showOptionDialog
        (
            SwingUtilities.getWindowAncestor(TextField),
            ErrorMessage,
            "Error",
            JOptionPane.YES_NO_OPTION,
            JOptionPane.ERROR_MESSAGE,
            null,
            new Object[] {"Edit", "Revert"},
            "Revert"
        );
	    
        if (answer == REVERT)
        {
        	TextField.setValue(TextField.getValue());
        	return true;
        }

        return false;
    }
	
	public boolean validEdit()
	{
    	final String text = TextField.getText();
    	
    	if ((text == null) || text.equals(""))
    		return false;
    	
    	final String textTrimmed = text.trim();

    	Number number;
    	final ParsePosition position = new ParsePosition(0);

        try
        {
        	number = FormatCurrencyUngrouped.parse(textTrimmed,position);
        }
        catch (Exception e)
        {
        	return false;
        }

        if ((number == null) || (position.getIndex() < textTrimmed.length()))
        	return false;

        final Double value = new Double(number.doubleValue());

        if ((value < MinimumValue) || (value > MaximumValue))
        	return false;
        
        if (text.length() != textTrimmed.length())
        	TextField.setText(textTrimmed);

        return true;
	}
	
	public void signalEdit(final boolean valid) throws ParseException
	{
		if (valid)
			TextField.commitEdit();

		TextField.postActionEvent();
	}
}