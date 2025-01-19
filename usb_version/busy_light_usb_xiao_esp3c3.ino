#include <Adafruit_NeoPixel.h>

#define PIN A0       // Pin where the NeoPixel is connected
#define NUMPIXELS 60 // Total number of pixels

Adafruit_NeoPixel strip = Adafruit_NeoPixel(NUMPIXELS, PIN, NEO_GRB + NEO_KHZ800);

// Function to light up the middle two rows
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

void setup() {
  Serial.begin(115200); // Initialize the serial communication
  strip.begin();
  strip.show(); // Initialize all pixels to 'off'

  // Welcome LED Blink (White)
  lightMiddleRows(strip.Color(100, 100, 100));
  delay(500);
  strip.clear();
  strip.show();
  delay(500);
}

void loop() {
  if (Serial.available() > 0) {
    char color = Serial.read();
    Serial.print("Received: ");
    Serial.println(color); // Echo back the received character

    delay(10); // Small delay to ensure the echo is properly transmitted

    // Apply LED changes based on the received character
    if (color == 'R') {
      // Red for microphone in use
      lightMiddleRows(strip.Color(100, 0, 0));
    } else if (color == 'G') {
      // Green for microphone not in use
      lightMiddleRows(strip.Color(0, 100, 0));
    } else if (color == 'W') {
      // White for no connection or lost connection
      lightMiddleRows(strip.Color(100, 100, 100));
    } else if (color == 'E') {
      // Turn off LEDs
      strip.clear();
      strip.show();
    }
  }
}
