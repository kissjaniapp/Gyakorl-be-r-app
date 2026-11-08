package hu.eredmenytablazat.app;

import android.os.Bundle;
import com.getcapacitor.BridgeActivity;

public class MainActivity extends BridgeActivity {
    @Override
    public void onCreate(Bundle savedInstanceState) {
        registerPlugin(JpegSaverPlugin.class);
        super.onCreate(savedInstanceState);
    }
}

