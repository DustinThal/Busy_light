#include <BLEDevice.h>
#include <BLEServer.h>
#include <BLEUtils.h>
#include <BLE2902.h>
#include <Adafruit_NeoPixel.h>

#define SERVICE_UUID        "12345678-1234-1234-1234-123456789abc"
#define CHARACTERISTIC_UUID "abcd1234-abcd-1234-abcd-12345678abcd"

#define PIN A0       // Pin where the NeoPixel is connected
#define NUMPIXELS 60 // Total number of pixels

Adafruit_NeoPixel strip = Adafruit_NeoPixel(NUMPIXELS, PIN, NEO_GRB + NEO_KHZ800);

// Connection timeout settings (in milliseconds)
const unsigned long MIN_CONNECTION_TIME = 5 * 60 * 1000; // 5 minutes
const unsigned long MAX_IDLE_TIME = 5 * 60 * 1000;       // 5 minutes of inactivity

BLEServer *server = nullptr;
BLECharacteristic *characteristic = nullptr;
unsigned long lastActivityTime = 0; // Tracks the last time data was received
bool deviceConnected = false;

// Function to set the LED matrix color
void lightMiddleRows(uint32_t color) {
  strip.clear(); // Clear previous colors
  for (int group = 0; group < 10; group++) {  // Iterate over 10 groups of 6
    int startRow = 6 * group;
    int mid1 = startRow + 2;
    int mid2 = startRow + 3;

    strip.setPixelColor(mid1, color);
    strip.setPixelColor(mid2, color);
  }
  strip.show(); // Apply the changes
}

// Custom server callbacks
class MyServerCallbacks : public BLEServerCallbacks {
  void onConnect(BLEServer* server) override {
    Serial.println("Device Connected.");
    deviceConnected = true;
    lastActivityTime = millis(); // Start the timer when the device connects
  }

  void onDisconnect(BLEServer* server) override {
    Serial.println("Device Disconnected.");
    deviceConnected = false;
    BLEDevice::startAdvertising(); // Restart advertising after disconnect
  }
};

// Custom security callbacks
class MySecurityCallbacks : public BLESecurityCallbacks {
  bool onConfirmPIN(uint32_t passkey) override {
    Serial.print("Numeric Comparison: ");
    Serial.println(passkey);
    return true; // Accept the numeric comparison
  }

  uint32_t onPassKeyRequest() override {
    Serial.println("PassKey Request Received.");
    return 123456; // Static passkey
  }

  void onPassKeyNotify(uint32_t passkey) override {
    Serial.print("PassKey Notify: ");
    Serial.println(passkey);
  }

  bool onSecurityRequest() override {
    Serial.println("Security Request Received. Allowing pairing.");
    return true;
  }

  void onAuthenticationComplete(esp_ble_auth_cmpl_t cmpl) override {
    if (cmpl.success) {
      Serial.println("Authentication Complete: Success.");
    } else {
      Serial.println("Authentication Failed.");
    }
  }
};

// Callback for characteristic write
class MyCharacteristicCallbacks : public BLECharacteristicCallbacks {
  void onWrite(BLECharacteristic *pCharacteristic) override {
    String value = pCharacteristic->getValue(); // Get the written value as Arduino String
    if (value.length() > 0) {
      Serial.print("Received command: ");
      Serial.println(value); // Log the received command

      // Parse the RGB values from the command
      int r, g, b;
      if (sscanf(value.c_str(), "%d,%d,%d", &r, &g, &b) == 3) {
        r = constrain(r, 0, 255);
        g = constrain(g, 0, 255);
        b = constrain(b, 0, 255);
        Serial.printf("Setting RGB to: %d, %d, %d\n", r, g, b);
        lightMiddleRows(strip.Color(r, g, b)); // Set the LED matrix color
      } else {
        Serial.println("Invalid command format. Expected: xxx,xxx,xxx");
      }

      lastActivityTime = millis(); // Reset the activity timer
    } else {
      Serial.println("Received an empty write.");
    }
  }
};

// Configure BLE security
void configureBLESecurity() {
  BLESecurity *security = new BLESecurity();
  security->setAuthenticationMode(ESP_LE_AUTH_REQ_SC_BOND); // Secure connection with bonding
  security->setCapability(ESP_IO_CAP_IO);                   // Numeric comparison capability
  security->setInitEncryptionKey(ESP_BLE_ENC_KEY_MASK | ESP_BLE_ID_KEY_MASK);

  // Set static passkey
  uint32_t passkey = 123456;
  esp_ble_gap_set_security_param(ESP_BLE_SM_SET_STATIC_PASSKEY, &passkey, sizeof(passkey));
  esp_ble_auth_req_t auth_req = ESP_LE_AUTH_REQ_SC_BOND;
  esp_ble_gap_set_security_param(ESP_BLE_SM_AUTHEN_REQ_MODE, &auth_req, sizeof(auth_req));
}

void setup() {
  Serial.begin(115200);
  Serial.println("Starting BLE...");

  // Initialize NeoPixel strip
  strip.begin();
  strip.show(); // Initialize all pixels to 'off'
  
  // Welcome LED Blink (White)
  lightMiddleRows(strip.Color(100, 100, 100));
  delay(500);
  strip.clear();
  strip.show();
  delay(500);

  // Initialize BLE Device
  BLEDevice::init("busy_light_2A1c");
  server = BLEDevice::createServer();

  // Set Server Callbacks
  server->setCallbacks(new MyServerCallbacks());

  // Define a BLE service and characteristic
  BLEService *service = server->createService(SERVICE_UUID);
  characteristic = service->createCharacteristic(
      CHARACTERISTIC_UUID,
      BLECharacteristic::PROPERTY_READ |
      BLECharacteristic::PROPERTY_WRITE
  );

  // Set Characteristic Callbacks
  characteristic->setCallbacks(new MyCharacteristicCallbacks());

  // Add a descriptor for notifications/indications (optional)
  characteristic->addDescriptor(new BLE2902());
  characteristic->setValue("Hello BLE!");

  // Start the service
  service->start();

  // Start advertising the service
  BLEAdvertising *advertising = BLEDevice::getAdvertising();
  advertising->addServiceUUID(SERVICE_UUID);
  advertising->setScanResponse(true);
  advertising->setMinPreferred(0x06);  // Minimum preferred connection interval
  advertising->setMaxPreferred(0x12); // Maximum preferred connection interval
  BLEDevice::startAdvertising();
  Serial.println("BLE advertising started...");

  // Configure BLE security
  configureBLESecurity();

  // Set Security Callbacks
  BLEDevice::setSecurityCallbacks(new MySecurityCallbacks());
}

void loop() {
  // Check connection timeout
  if (deviceConnected) {
    unsigned long currentTime = millis();
    if ((currentTime - lastActivityTime) >= MAX_IDLE_TIME) {
      Serial.println("No activity detected. Disconnecting device.");
      server->disconnect(0); // Disconnect the client
      deviceConnected = false;
    }
  }

  delay(1000); // Keep the BLE server running
}
