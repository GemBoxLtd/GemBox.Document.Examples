# Execute docker commands from project's folder

```cmd
:: 1. Create docker image named "gemboximage".
docker build -t gemboximage -f Dockerfile .

:: 2. Create and start docker container named "gemboxcontainer".
docker run --name gemboxcontainer gemboximage

:: 3. Copy output files from container to project's folder.
docker cp gemboxcontainer:/app/output.docx .
docker cp gemboxcontainer:/app/output.pdf .

:: 4. Clean up, remove container and image.
docker image rm gemboximage
docker container rm gemboxcontainer
```