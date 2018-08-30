package baseltools;

import java.awt.Color;
import java.awt.Dimension;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;

import javax.swing.RepaintManager;
import javax.swing.border.MatteBorder;

public final class Environment
{
	private static final transient DecimalFormatSymbols DecimalFormatSymbols;

	public static final Color ColorBeige;
	public static final Color ColorBlue;
	public static final Color ColorBorder;
	public static final Color ColorBrown;
	public static final Color ColorCyan;
	public static final Color ColorDarkBlue;
	public static final Color ColorDarkGreen;
	public static final Color ColorDarkYellow;
	public static final Color ColorDisabled;
	public static final Color ColorEnabled;
	public static final Color ColorFont;
	public static final Color ColorGray;
	public static final Color ColorGreen;
	public static final Color ColorHeader;
	public static final Color ColorIndigo;
	public static final Color ColorLightBlue;
	public static final Color ColorLightGreen;
	public static final Color ColorLightYellow;
	public static final Color ColorMagenta;
	public static final Color ColorOrange;
	public static final Color ColorPink;
	public static final Color ColorPurple;
	public static final Color ColorRed;
	public static final Color ColorTeal;
	public static final Color ColorYellow;
	public static final Color[] ColorMap;
	public static final DecimalFormat FormatCurrencyGrouped;
	public static final DecimalFormat FormatCurrencyUngrouped;
	public static final HyperlinkBrowser HyperlinkBrowser;
	public static final MatteBorder BorderHeaderColumn;
	public static final MatteBorder BorderHeaderRow;
	public static final SimpleDateFormat DateFormat;

	public static Color ColorCcrColsICA;
	public static Color ColorCcrColsVM;
	public static Color ColorCcrNetsStd;
	public static Color ColorCcrNetsTrd;
	public static Color ColorCcrTrdsCo;
	public static Color ColorCcrTrdsCr;
	public static Color ColorCcrTrdsEq;
	public static Color ColorCcrTrdsFx;
	public static Color ColorCcrTrdsIr;
	public static Color ColorCcrTrdsRe;

	public static Color ColorOpBusILDC;
	public static Color ColorOpBusFC;
	public static Color ColorOpBusSC;
	public static Color ColorOpBusUBI;
	public static Color ColorOpBusBI;
	public static Color ColorOpBusBIC;
	public static Color ColorOpCapAll;
	public static Color ColorOpCapB2;
	public static Color ColorOpCapSMA;
	public static Color ColorOpLossAll;
	public static Color ColorOpLoss10;
	public static Color ColorOpLoss100;
	public static Color ColorOpLossLC;

	static
	{
		DecimalFormatSymbols = new DecimalFormatSymbols();
		DecimalFormatSymbols.setDecimalSeparator(',');
		DecimalFormatSymbols.setGroupingSeparator('.');
		
		ColorBeige = new Color(214, 201, 174);			// #D6C9AE
		ColorBlue = new Color(128, 177, 211);			// #80B1D3
		ColorBorder = Color.BLACK;
		ColorBrown = new Color(201, 149, 121);			// #C99579
		ColorCyan = new Color(191, 242, 242);			// #BFF2F2
		ColorDarkBlue = new Color(129, 142, 236);		// #818EEC
		ColorDarkGreen = new Color(110, 219, 103);		// #6EDB67
		ColorDarkYellow = new Color(223, 187, 78);		// #DFBB4E
		ColorDisabled = new Color(240, 240, 240);		// #F0F0F0
		ColorEnabled = Color.WHITE;
		ColorFont = Color.BLACK;
		ColorGray = new Color(164, 164, 164);			// #A4A4A4
		ColorGreen = new Color(154, 197, 80);			// #9AC550
		ColorHeader = new Color(246, 247, 249);			// #F6F7F9
		ColorIndigo = new Color(190, 186, 218);			// #BEBADA
		ColorLightBlue = new Color(141, 211, 199);		// #8DD3C7
		ColorLightGreen = new Color(206, 238, 88);		// #CEEE58
		ColorLightYellow = new Color(247, 247, 171);	// #F7F7AB
		ColorMagenta = new Color(227, 104, 125);		// #E3687D
		ColorOrange = new Color(253, 180, 98);			// #FDB462
		ColorPink = new Color(252, 205, 229);			// #FCCDE5
		ColorPurple = new Color(188, 128, 189);			// #BC80BD
		ColorRed = new Color(226, 103, 89);				// #E26759
		ColorTeal = new Color(189, 220, 182);			// #BDDCB6
		ColorYellow = new Color(255, 237, 111);			// #FFED6F

		ColorMap = new Color[]
		{
			ColorTeal,
			ColorOrange,
			ColorDarkGreen,
			ColorRed,
			ColorBlue,
			ColorYellow,
			ColorBrown,
			ColorPurple,
			ColorGray,
			ColorPink,
			ColorDarkBlue,
			ColorLightYellow,
			ColorLightBlue,
			ColorGreen,
			ColorMagenta,
			ColorIndigo,
			ColorDarkYellow,
			ColorBeige,
			ColorLightGreen,
			ColorCyan
		};

		FormatCurrencyGrouped = new DecimalFormat();
		FormatCurrencyGrouped.setGroupingUsed(true);
		FormatCurrencyGrouped.setDecimalFormatSymbols(DecimalFormatSymbols);
		FormatCurrencyGrouped.setMaximumFractionDigits(2);
		FormatCurrencyGrouped.setMinimumFractionDigits(2);	

		FormatCurrencyUngrouped = new DecimalFormat();
		FormatCurrencyUngrouped.setGroupingUsed(false);
		FormatCurrencyUngrouped.setDecimalFormatSymbols(DecimalFormatSymbols);
		FormatCurrencyUngrouped.setMaximumFractionDigits(2);
		FormatCurrencyUngrouped.setMinimumFractionDigits(2);

		HyperlinkBrowser = new HyperlinkBrowser();
		
		BorderHeaderColumn = new MatteBorder(1,0,1,0,ColorBorder);
		BorderHeaderRow = new MatteBorder(0,0,0,1,ColorBorder);
		
		DateFormat = new SimpleDateFormat("dd/MM/yyyy");
		DateFormat.setLenient(false);
		
		ColorCcrColsICA = ColorPurple;
		ColorCcrColsVM = ColorIndigo;
		ColorCcrNetsStd = ColorBrown;
		ColorCcrNetsTrd = ColorBeige;
		ColorCcrTrdsCo = ColorOrange;
		ColorCcrTrdsCr = ColorGreen;
		ColorCcrTrdsEq = ColorBlue;
		ColorCcrTrdsFx = ColorYellow;
		ColorCcrTrdsIr = ColorMagenta;
		ColorCcrTrdsRe = Color.WHITE;

		ColorOpBusILDC = ColorTeal;
		ColorOpBusFC = ColorYellow;
		ColorOpBusSC = ColorIndigo;
		ColorOpBusUBI = ColorBeige;
		ColorOpBusBI = ColorBrown;
		ColorOpBusBIC = ColorRed;
		ColorOpCapAll = ColorPurple;
		ColorOpCapB2 = ColorDarkYellow;
		ColorOpCapSMA = ColorMagenta;
		ColorOpLossAll = ColorGreen;
		ColorOpLoss10 = ColorOrange;
		ColorOpLoss100 = ColorRed;
		ColorOpLossLC = ColorDarkBlue;
	}

    private Environment()
    {
        throw new AssertionError();
    }

    public static String FormatNumber(final Object value, final boolean grouping, final int digits)
    {
		final DecimalFormat numberFormat = new DecimalFormat();
		numberFormat.setGroupingUsed(grouping);
		numberFormat.setDecimalFormatSymbols(DecimalFormatSymbols);
		numberFormat.setMaximumFractionDigits(digits);
		numberFormat.setMinimumFractionDigits(digits);
		
		return numberFormat.format(value);
    }
    
    public static void CleanMemoryHeap()
    {
    	try
    	{
	    	final RepaintManager rm = RepaintManager.currentManager(null);
	    	final Dimension localDimension = rm.getDoubleBufferMaximumSize();
	    	
	    	rm.setDoubleBufferMaximumSize(new Dimension(0, 0));
	    	rm.setDoubleBufferMaximumSize(localDimension);
	
	    	System.gc();
    	}
    	catch (Exception e) { }
    }
}