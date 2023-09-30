# Image to Excel Converter

This utility allows you to convert a images to an Excel sheet, where the background color of each cell corresponds to a pixel's RGB value from the image. Additionally, it comes dockerized, enabling a simple and environment-independent execution. The intended use case for this tool is easily converts images to knitting patterns. 

## Getting Started

### Prerequisites

- Docker installed on your system. If not, get it from [Docker's official site](https://docs.docker.com/get-docker/).

### Building the Docker Image

1. Clone this repository:
   ```bash
   git clone https://github.com/Aramees556/img-to-excel
   cd img-to-excel
   ```

2. Build the Docker image:
   ```bash
   docker build -t png-to-excel .
   ```

### Running the Utility

1. To execute the script, use the following command:
   ```bash
   docker run -v /path/to/your/files:/app img-to-excel
   ```

   - Replace `/path/to/your/files` with the path where your images and desired output Excel files are located or will be saved.
   
   This will launch the script and prompt you for inputs. After processing, the Excel files will be saved to the specified path on your host system.

## How It Works

The script first prompts you for the path to the PNG image. Then, it asks for the desired grid width and height, which determines the size of the image representation in Excel. Finally, the path to save the output Excel file is requested. Each cell in the Excel file will have a background color corresponding to the RGB value of the respective pixel from the resized image.

## License

This project is open-sourced software licensed under the MIT license. 

---
