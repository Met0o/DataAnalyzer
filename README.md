# PorAnalyzer

https://gist.github.com/roaldnefs/fe9f36b0e8cf2890af14572c083b516c

https://www.xquartz.org/

X Server Setup (XQuartz for macOS):

To use a graphical application from a Docker container, you could use an X Server.
On macOS, you could install XQuartz, which allows macOS to run X11 applications.
You'd then need to configure Docker to forward the display to XQuartz:
Run XQuartz on macOS.
In a terminal on macOS, allow access to the X server by running:
bash
Copy code
xhost + 127.0.0.1
When running the container, set the DISPLAY environment variable to point to the macOS host:
bash
Copy code
docker run -e DISPLAY=host.docker.internal:0 -v /path/to/local/folder:/data your-docker-image