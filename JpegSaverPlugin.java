package hu.eredmenytablazat.app;

import android.content.ContentResolver;
import android.content.ActivityNotFoundException;
import android.content.ContentValues;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Environment;
import android.provider.MediaStore;
import android.util.Base64;

import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import androidx.core.content.FileProvider;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;

import org.json.JSONObject;

@CapacitorPlugin(name = "JpegSaver")
public class JpegSaverPlugin extends Plugin {

    @PluginMethod
    public void save(PluginCall call) {
        String data = call.getString("data");
        String name = call.getString("name", "eredmeny.jpg");
        if (data == null) {
            call.reject("Nincs adat");
            return;
        }
        try {
            byte[] bytes = Base64.decode(data, Base64.DEFAULT);
            ContentResolver resolver = getContext().getContentResolver();
            ContentValues values = new ContentValues();
            values.put(MediaStore.Images.Media.DISPLAY_NAME, name);
            values.put(MediaStore.Images.Media.MIME_TYPE, "image/jpeg");
            if (Build.VERSION.SDK_INT >= 29) {
                values.put(MediaStore.Images.Media.RELATIVE_PATH, Environment.DIRECTORY_PICTURES + "/Eredmenytablazat");
            }
            Uri uri = resolver.insert(MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values);
            if (uri == null) {
                call.reject("Nem sikerült létrehozni a fájlt");
                return;
            }
            OutputStream out = resolver.openOutputStream(uri);
            if (out == null) {
                call.reject("Nem sikerült megnyitni a fájlt");
                return;
            }
            out.write(bytes);
            out.close();
            JSObject ret = new JSObject();
            ret.put("uri", uri.toString());
            call.resolve(ret);
        } catch (Exception e) {
            call.reject("Mentési hiba: " + e.getMessage());
        }
    }

    @PluginMethod
    public void sendTo(PluginCall call) {
        String data = call.getString("data");
        String name = call.getString("name", "eredmeny.jpg");
        String pkg = call.getString("pkg");
        if (data == null || pkg == null) {
            call.reject("Hiányzó adat");
            return;
        }
        try {
            byte[] bytes = Base64.decode(data, Base64.DEFAULT);
            File f = new File(getContext().getCacheDir(), name);
            FileOutputStream fos = new FileOutputStream(f);
            fos.write(bytes);
            fos.close();
            Uri uri = FileProvider.getUriForFile(getContext(), getContext().getPackageName() + ".fileprovider", f);
            Intent i = new Intent(Intent.ACTION_SEND);
            i.setType("image/jpeg");
            i.putExtra(Intent.EXTRA_STREAM, uri);
            i.setPackage(pkg);
            i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_ACTIVITY_NEW_TASK);
            getContext().startActivity(i);
            call.resolve();
        } catch (ActivityNotFoundException e) {
            call.reject("Az app nincs telepítve");
        } catch (Exception e) {
            call.reject("Küldési hiba: " + e.getMessage());
        }
    }

    @PluginMethod
    public void sendMany(PluginCall call) {
        String pkg = call.getString("pkg");
        JSArray files = call.getArray("files");
        if (pkg == null || files == null || files.length() == 0) {
            call.reject("Hiányzó adat");
            return;
        }
        try {
            ArrayList<Uri> uris = new ArrayList<>();
            for (int k = 0; k < files.length(); k++) {
                JSONObject o = files.getJSONObject(k);
                byte[] bytes = Base64.decode(o.getString("data"), Base64.DEFAULT);
                File f = new File(getContext().getCacheDir(), o.getString("name"));
                FileOutputStream fos = new FileOutputStream(f);
                fos.write(bytes);
                fos.close();
                uris.add(FileProvider.getUriForFile(getContext(), getContext().getPackageName() + ".fileprovider", f));
            }
            Intent i = new Intent(uris.size() > 1 ? Intent.ACTION_SEND_MULTIPLE : Intent.ACTION_SEND);
            i.setType("image/jpeg");
            if (uris.size() > 1) {
                i.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
            } else {
                i.putExtra(Intent.EXTRA_STREAM, uris.get(0));
            }
            i.setPackage(pkg);
            i.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_ACTIVITY_NEW_TASK);
            getContext().startActivity(i);
            call.resolve();
        } catch (ActivityNotFoundException e) {
            call.reject("Az app nincs telepítve");
        } catch (Exception e) {
            call.reject("Küldési hiba: " + e.getMessage());
        }
    }

    @PluginMethod
    public void saveFile(PluginCall call) {
        String data = call.getString("data");
        String name = call.getString("name", "eredmeny.xlsx");
        String mime = call.getString("mime", "application/octet-stream");
        if (data == null) {
            call.reject("Nincs adat");
            return;
        }
        if (Build.VERSION.SDK_INT < 29) {
            call.reject("Ehhez Android 10 vagy újabb kell");
            return;
        }
        try {
            byte[] bytes = Base64.decode(data, Base64.DEFAULT);
            ContentResolver resolver = getContext().getContentResolver();
            ContentValues values = new ContentValues();
            values.put(MediaStore.MediaColumns.DISPLAY_NAME, name);
            values.put(MediaStore.MediaColumns.MIME_TYPE, mime);
            values.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS + "/Eredmenytablazat");
            Uri uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
            if (uri == null) {
                call.reject("Nem sikerült létrehozni a fájlt");
                return;
            }
            OutputStream out = resolver.openOutputStream(uri);
            if (out == null) {
                call.reject("Nem sikerült megnyitni a fájlt");
                return;
            }
            out.write(bytes);
            out.close();
            JSObject ret = new JSObject();
            ret.put("uri", uri.toString());
            call.resolve(ret);
        } catch (Exception e) {
            call.reject("Mentési hiba: " + e.getMessage());
        }
    }
}
