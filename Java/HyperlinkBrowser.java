package baseltools;

import java.awt.Desktop;
import javax.swing.event.HyperlinkEvent;
import javax.swing.event.HyperlinkListener;

public final class HyperlinkBrowser implements HyperlinkListener
{
    @Override
    public void hyperlinkUpdate(final HyperlinkEvent event)
    {
    	if (event.getEventType() != HyperlinkEvent.EventType.ACTIVATED)
    		return;
    	
    	if (!Desktop.isDesktopSupported())
    		return;

        try
        {
        	Desktop.getDesktop().browse(event.getURL().toURI());
        }
        catch (Exception e) { }
    }
}