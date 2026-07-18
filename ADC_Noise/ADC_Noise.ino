const int analogPin1 = A0;
const int analogPin2 = A1;
const int N = 256;              // FFTなら1024点がおすすめ

uint16_t data1[N];
uint16_t data2[N];

void setup() {
  Serial.begin(115200);

  Serial.println("READY");
}

void loop() {

  if (Serial.available()) {

    String cmd = Serial.readStringUntil('\n');
    cmd.trim();

    if (cmd == "START") {

      unsigned long tStart = micros();

      for (int i = 0; i < N; i++) {
        data1[i] = analogRead(analogPin1);
        data2[i] = analogRead(analogPin2);
        
        // 必要ならサンプリング周期を調整
        // delayMicroseconds(100);
      }

      unsigned long tEnd = micros();

      // ===== サンプリング周波数計算 =====
      float elapsed = (tEnd - tStart) / 1000000.0;   // 秒
      float Fs = (N - 1) / elapsed;                  // Hz

      Serial.println("BEGIN");
      Serial.print("Fs=");
      Serial.println(Fs, 2);

      for (int i = 0; i < N; i++) {
        Serial.print(data1[i]);
        Serial.print("\t");
        Serial.println(data2[i]);
      }

      Serial.println("END");
    }
  }
}
