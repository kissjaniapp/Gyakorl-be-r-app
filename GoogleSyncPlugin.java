package hu.eredmenytablazat.app;

import android.content.Context;
import androidx.annotation.NonNull;
import androidx.work.BackoffPolicy;
import androidx.work.Constraints;
import androidx.work.Data;
import androidx.work.ExistingWorkPolicy;
import androidx.work.NetworkType;
import androidx.work.OneTimeWorkRequest;
import androidx.work.WorkManager;
import androidx.work.Worker;
import androidx.work.WorkerParameters;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;
import org.json.JSONObject;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.util.concurrent.TimeUnit;

@CapacitorPlugin(name = "GoogleSync")
public class GoogleSyncPlugin extends Plugin {
    private static final String FILE_PREFIX = "sheet_pending_";

    @PluginMethod
    public void configure(PluginCall call) {
        try {
            String url = call.getString("url", "");
            String secret = call.getString("secret", "");
            if ((!url.isEmpty() && !isAppsScriptEndpoint(url)) || url.length() > 4096 || secret.length() > 4096) {
                call.reject("Invalid Google synchronization settings."); return;
            }
            android.content.SharedPreferences prefs = getContext().getSharedPreferences("google_sync_config", Context.MODE_PRIVATE);
            if (!url.equals(prefs.getString("url", null)) || !secret.equals(prefs.getString("secret", null))) {
                if (!prefs.edit().putString("url", url).putString("secret", secret).commit())
                    throw new java.io.IOException("Could not store configuration");
            }
            JSObject result = new JSObject(); result.put("ok", true); call.resolve(result);
        } catch (Exception ex) { call.reject("Could not store settings.", ex); }
    }

    @PluginMethod
    public void enqueue(PluginCall call) {
        try {
            String url = call.getString("url", "");
            String secret = call.getString("secret", "");
            JSObject payload = call.getObject("payload");
            if (payload == null || secret.isEmpty() || !isAppsScriptEndpoint(url)) {
                call.reject("Invalid Google synchronization settings.");
                return;
            }
            String eventId = payload.optString("eventId", "");
            if (eventId.isEmpty() || eventId.length() > 256) {
                call.reject("Invalid event identifier.");
                return;
            }
            JSONObject stored = new JSONObject();
            stored.put("url", url);
            stored.put("secret", secret);
            stored.put("payload", payload);
            stored.put("revision", java.util.UUID.randomUUID().toString());
            String id = hash(eventId);
            File target = new File(getContext().getFilesDir(), FILE_PREFIX + id + ".json");
            // File names are derived from SHA-256; do not use event names as paths.
            synchronized (GoogleSyncPlugin.class) {
                File temp = File.createTempFile(FILE_PREFIX, ".tmp", getContext().getFilesDir());
                try {
                    try (FileOutputStream out = new FileOutputStream(temp)) {
                        out.write(stored.toString().getBytes(StandardCharsets.UTF_8));
                        out.getFD().sync();
                    }
                    // Android's rename replaces atomically; if it fails, keep old file intact.
                    if (!temp.renameTo(target)) throw new java.io.IOException("Could not store pending result");
                } finally { if (temp.exists()) temp.delete(); }
                Data input = new Data.Builder().putString("file", target.getName()).build();
                Constraints constraints = new Constraints.Builder()
                    .setRequiredNetworkType(NetworkType.CONNECTED).build();
                OneTimeWorkRequest request = new OneTimeWorkRequest.Builder(SyncWorker.class)
                    .setInputData(input).setConstraints(constraints)
                    .setBackoffCriteria(BackoffPolicy.EXPONENTIAL, 30, TimeUnit.SECONDS).build();
                // Wait until WorkManager has durably saved the request before acknowledging JS.
                WorkManager.getInstance(getContext()).enqueueUniqueWork(
                    "sheet-sync-" + id, ExistingWorkPolicy.APPEND_OR_REPLACE, request)
                    .getResult().get(15, TimeUnit.SECONDS);
            }
            JSObject result = new JSObject(); result.put("ok", true); call.resolve(result);
        } catch (Exception ex) {
            call.reject("Could not queue result.", ex);
        }
    }

    private static boolean isAppsScriptEndpoint(String value) {
        try {
            URL u = new URL(value);
            return "https".equalsIgnoreCase(u.getProtocol()) &&
                "script.google.com".equalsIgnoreCase(u.getHost()) &&
                u.getPort() == -1 && u.getUserInfo() == null &&
                u.getPath().startsWith("/macros/s/") && u.getPath().endsWith("/exec");
        } catch (Exception ex) { return false; }
    }
    private static boolean isTrustedRedirect(URL u) {
        String host = u.getHost().toLowerCase(java.util.Locale.ROOT);
        return "https".equalsIgnoreCase(u.getProtocol()) && u.getPort() == -1 &&
            u.getUserInfo() == null && (host.equals("script.google.com") ||
            host.equals("script.googleusercontent.com"));
    }
    private static String hash(String s) throws Exception {
        byte[] bytes = MessageDigest.getInstance("SHA-256").digest(s.getBytes(StandardCharsets.UTF_8));
        StringBuilder b = new StringBuilder(); for (byte x : bytes) b.append(String.format("%02x", x & 255));
        return b.toString();
    }

    public static class SyncWorker extends Worker {
        public SyncWorker(@NonNull Context context, @NonNull WorkerParameters params) { super(context, params); }
        @NonNull @Override public Result doWork() {
            String name = getInputData().getString("file");
            if (name == null || !name.matches("sheet_pending_[0-9a-f]{64}\\.json")) return Result.failure();
            File file = new File(getApplicationContext().getFilesDir(), name);
            try {
                JSONObject job;
                synchronized (GoogleSyncPlugin.class) {
                    if (!file.exists()) return Result.success();
                    job = readJob(file);
                }
                String revision = job.getString("revision");
                android.content.SharedPreferences prefs = getApplicationContext().getSharedPreferences("google_sync_config", Context.MODE_PRIVATE);
                String endpoint = prefs.getString("url", job.getString("url"));
                String secret = prefs.getString("secret", job.getString("secret"));
                if (secret == null || secret.isEmpty() || !isAppsScriptEndpoint(endpoint)) return Result.retry();
                JSONObject payload=job.getJSONObject("payload");
                String eventId=payload.getString("eventId");
                payload.put("secret",secret);
                if (!postAndVerify(endpoint,payload,eventId)) return Result.retry();
                synchronized (GoogleSyncPlugin.class) {
                    if (!file.exists()) return Result.success();
                    // A newer edit to this competition must also reach the server.
                    if (!revision.equals(readJob(file).optString("revision"))) return Result.retry();
                    if (!file.delete() && file.exists()) return Result.retry();
                }
                return Result.success();
            } catch (Exception ex) { return Result.retry(); }
        }
        private JSONObject readJob(File file) throws Exception {
            ByteArrayOutputStream body = new ByteArrayOutputStream();
            try (InputStream in = new FileInputStream(file)) {
                byte[] b = new byte[4096]; int n;
                while ((n=in.read(b))!=-1) {
                    body.write(b,0,n);
                    if(body.size()>2_000_000) throw new java.io.IOException("Job too large");
                }
            }
            return new JSONObject(new String(body.toByteArray(), StandardCharsets.UTF_8));
        }
        private boolean postAndVerify(String endpoint, JSONObject payload, String eventId) throws Exception {
            URL url = new URL(endpoint);
            boolean post = true;
            byte[] content = payload.toString().getBytes(StandardCharsets.UTF_8);
            for (int i=0; i<6; i++) {
                if (!isTrustedRedirect(url)) return false;
                HttpURLConnection c = (HttpURLConnection) url.openConnection();
                try {
                    c.setInstanceFollowRedirects(false);
                    c.setConnectTimeout(12000); c.setReadTimeout(15000);
                    c.setRequestMethod(post ? "POST" : "GET");
                    if (post) {
                        c.setDoOutput(true);
                        c.setRequestProperty("Content-Type", "text/plain; charset=UTF-8");
                        c.setFixedLengthStreamingMode(content.length);
                        try (OutputStream out=c.getOutputStream()) {out.write(content);}
                    }
                    int status=c.getResponseCode();
                    if (status==301 || status==302 || status==303 || status==307 || status==308) {
                        String location=c.getHeaderField("Location");
                        if (location==null) return false;
                        url=new URL(url,location);
                        if(status==301 || status==302 || status==303) post=false;
                        continue;
                    }
                    if (status<200 || status>=300) return false;
                    ByteArrayOutputStream bytes=new ByteArrayOutputStream();
                    try (InputStream in=c.getInputStream()) {
                        byte[] buf=new byte[4096]; int n;
                        while((n=in.read(buf))!=-1) {bytes.write(buf,0,n); if(bytes.size()>65536)return false;}
                    }
                    JSONObject response=new JSONObject(new String(bytes.toByteArray(), StandardCharsets.UTF_8));
                    return response.optBoolean("ok",false) && eventId.equals(response.optString("eventId"));
                } finally {c.disconnect();}
            }
            return false;
        }
    }
}
