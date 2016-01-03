# https://www.realvnc.com/products/vnc/raspberrypi/

curl -L -o VNC.tar.gz https://www.realvnc.com/download/binary/latest/debian/arm/
tar xvf VNC.tar.gz
sudo dpkg -i <VNC-Server-package-name>.deb <VNC-Viewer-package-name>.deb
sudo vnclicense -add <license key>

# Service Mode

sudo vncpasswd /root/.vnc/config.d/vncserver-x11

#Debian 8 'Jessie'
sudo systemctl start vncserver-x11-serviced.service
#Debian 7 'Wheezy'
sudo /etc/init.d/vncserver-x11-serviced start

#Debian 8 'Jessie'
sudo systemctl enable vncserver-x11-serviced.service
#Debian 7 'Wheezy'
sudo update-rc.d vncserver-x11-serviced defaults

#RealVNC Viewer downloads - https://www.realvnc.com/download/viewer/windows/
