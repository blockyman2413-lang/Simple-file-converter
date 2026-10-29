// Test fixture: deliberately stalls so the supervisor's timeout can be verified.
class SlowEngine { static void Main() { System.Threading.Thread.Sleep(30000); } }
