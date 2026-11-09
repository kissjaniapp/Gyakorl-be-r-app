package hu.eredmenytablazat.app;

import android.content.Context;
import android.net.ConnectivityManager;
import android.net.Network;
import android.net.NetworkCapabilities;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

@CapacitorPlugin(name = "ArduinoSensor")
public class ArduinoSensorPlugin extends Plugin {
    private final ExecutorService executor = Executors.newSingleThreadExecutor();

    @PluginMethod
    public void get(PluginCall call) {
        final String path = call.getString("path", "");
        if (!path.matches("/(?:status|arm\\?token=[A-Za-z0-9_%-]+(?:&previous=[A-Za-z0-9_%-]+)?)")) {
            call.reject("Invalid sensor request");
            return;
        }
        executor.execute(() -> {
            ConnectivityManager manager = (ConnectivityManager) getContext().getSystemService(Context.CONNECTIVITY_SERVICE);
            if (manager == null) { call.reject("Wi-Fi network unavailable"); return; }
            try {
                for (Network network : manager.getAllNetworks()) {
                    NetworkCapabilities caps = manager.getNetworkCapabilities(network);
                    if (caps == null || !caps.hasTransport(NetworkCapabilities.TRANSPORT_WIFI)) continue;
                    HttpURLConnection connection = null;
                    try {
                        connection = (HttpURLConnection) network.openConnection(new URL("http://192.168.4.1" + path));
                        connection.setConnectTimeout(2500); connection.setReadTimeout(2500);
                        connection.setRequestMethod("GET"); connection.setUseCaches(false);
                        int status = connection.getResponseCode();
                        InputStream stream = status >= 400 ? connection.getErrorStream() : connection.getInputStream();
                        if (stream == null) continue;
                        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
                        try (InputStream input = stream) {
                            byte[] buffer = new byte[1024]; int count;
                            while ((count = input.read(buffer)) != -1) {
                                bytes.write(buffer, 0, count);
                                if (bytes.size() > 16384) throw new java.io.IOException("Sensor response too large");
                            }
                        }
                        JSObject data = new JSObject(new String(bytes.toByteArray(), StandardCharsets.UTF_8));
                        // Avoid accepting a different device that happens to use the same address.
                        if (!SensorResponseValidator.accepts(status, path, data.optString("boot_id", ""), data.optString("state", ""), data.optString("error", ""))) continue;
                        JSObject result = new JSObject();
                        result.put("status", status); result.put("data", data); call.resolve(result); return;
                    } catch (Exception ignored) {
                        // A different Wi-Fi network can be listed; try the next one.
                    } finally { if (connection != null) connection.disconnect(); }
                }
                call.reject("Időmérő kártya nem érhető el Wi-Fi kapcsolaton");
            } catch (Exception ex) { call.reject("A Wi-Fi kapcsolat nem érhető el", ex); }
        });
    }
}
