git clone https://github.com/DexterInd/BrickPi.git
cd "/home/pi/BrickPi/Setup Files/"
sudo chmod +x install.sh
sudo ./install.sh
cd "/home/pi/BrickPi/Setup Files/wiringPi"
sudo chmod 777 build
./build
sudo apt-get install libi2c-dev
gpio load i2c 10
sudo apt-get install python-serial
sudo apt-get install python-rpi.gpio
sudo apt-get install i2c-tools
