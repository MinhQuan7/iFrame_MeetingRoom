#define ERA_DEBUG
#define ERA_SERIAL Serial
#define ERA_LOCATION_VN
// #define ERA_LOCATION_SG
// You should get Auth Token in the ERa App or ERa Dashboard
#define ERA_AUTH_TOKEN "ca614c19-2a5c-4a12-9cd8-895cf754d84e"

#include <Arduino.h>
#include <ERa.hpp>

const char ssid[] = "eoh.io";
const char pass[] = "Eoh@2020";

ERA_CONNECTED()
{
  ERA_LOG("ERa", "ERa connected!");
}

/* This function will run every time ERa is disconnected */
ERA_DISCONNECTED()
{
  ERA_LOG("ERa", "ERa disconnected!");
}

/* This function send uptime every second to Virtual Pin 1 */
void timerEvent()
{
  ERA_LOG("Timer", "Uptime: %d", ERaMillis() / 1000L);
  // Tạo giá trị ngẫu nhiên cho dòng điện từ 0.5 đến 10 Amps
  float current = random(5, 101) / 10.0; // Tạo số nguyên từ 5 đến 100, chia cho 10 để có giá trị từ 0.5 đến 10.0
  ERa.virtualWrite(V0, current);
  // Tạo giá trị ngẫu nhiên cho điện áp từ 220 đến 240 Volts
  float voltage = random(2200, 2401) / 10.0; // Tạo số nguyên từ 2200 đến 2400, chia cho 10 để có giá trị từ 220.0 đến 240.0
  ERa.virtualWrite(V1, voltage);
  // Tính công suất tiêu thụ (P = V * I)
  float power_consumption = current * voltage;
  ERa.virtualWrite(V2, power_consumption);
  // In ra các giá trị lên Serial Monitor
  Serial.printf("Current: %.2f A, Voltage: %.2f V, Power Consumption: %.2f W\n", current, voltage, power_consumption);

  delay(2000); // Đợi 2 giây trước khi lặp lại
}

void setup()
{
  /* Setup debug console */
#if defined(ERA_DEBUG)
  Serial.begin(115200);
#endif
  ERa.setScanWiFi(true);
  /* Initializing the ERa library. */
  ERa.begin(ssid, pass);
  /* Setup timer called function every second */
  ERa.addInterval(1000L, timerEvent);
}

// Hàm vòng lặp chính
void loop()
{
  ERa.run();
}