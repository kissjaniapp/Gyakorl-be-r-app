package hu.eredmenytablazat.app;

/** Validates expected timer responses before accepting data from a local IP. */
public final class SensorResponseValidator {
    private SensorResponseValidator() {}
    public static boolean accepts(int status, String path, String boot, String state, String error) {
        if (status >= 200 && status < 300) {
            return !boot.isEmpty() && (state.equals("idle") || state.equals("armed") ||
                state.equals("running") || state.equals("finished") || state.equals("error"));
        }
        return status == 409 && path.startsWith("/arm?token=") &&
            (error.equals("measurement_running") || error.equals("previous_not_confirmed"));
    }
}
