package hu.eredmenytablazat.app;
public class SensorResponseValidatorTest {
  public static void main(String[] args) {
    String p="/arm?token=abc";
    check(SensorResponseValidator.accepts(200,"/status","board","running",""));
    check(SensorResponseValidator.accepts(409,p,"","","measurement_running"));
    check(SensorResponseValidator.accepts(409,p,"","","previous_not_confirmed"));
    check(!SensorResponseValidator.accepts(409,"/status","","","measurement_running"));
    check(!SensorResponseValidator.accepts(409,p,"","","unrelated"));
    check(!SensorResponseValidator.accepts(200,"/status","","idle",""));
    check(!SensorResponseValidator.accepts(200,"/status","board","unexpected",""));
    System.out.println("Android timer response validation passed");
  }
  static void check(boolean value) { if (!value) throw new AssertionError(); }
}
