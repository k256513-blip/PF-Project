#define VIB_PIN     2
#define BUTTON_PIN  7
#define LED_PIN     8
#define BUZZER_PIN  9

bool alarmRunning = false;

void setup() {
    pinMode(VIB_PIN,    INPUT);
    pinMode(BUTTON_PIN, INPUT_PULLUP);
    pinMode(LED_PIN,    OUTPUT);
    pinMode(BUZZER_PIN, OUTPUT);
    Serial.begin(9600); 
}

void loop() {
    if (digitalRead(VIB_PIN) == LOW && !alarmRunning) {
        alarmRunning = true;
        Serial.print("THEFT_DETECTED\n");
        Serial.flush(); 

        unsigned long alarmStart = millis();
        bool manualStop = false;

        // Rapid Fire Alarm for 15 seconds
        while (millis() - alarmStart < 15000) {
            if (digitalRead(BUTTON_PIN) == LOW) {
                manualStop = true;
                break;
            }
            digitalWrite(LED_PIN, HIGH); digitalWrite(BUZZER_PIN, HIGH);
            delay(50);
            digitalWrite(LED_PIN, LOW);  digitalWrite(BUZZER_PIN, LOW);
            delay(50);
        }

        if (manualStop) {
            stopAlarm("STOPPED_BY_BUTTON");
        } else {
            stopAlarm("ALARM_TIMEOUT");
        }
    }
}

void stopAlarm(String reason) {
    alarmRunning = false;
    digitalWrite(LED_PIN,    LOW);
    digitalWrite(BUZZER_PIN, LOW);
    Serial.println(reason); // Sends the specific reason to C
    Serial
