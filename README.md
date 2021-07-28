# Powerpoint Image Source finder

A python script for finding missing image web-sources in pptx files.
The script will parse each slide of a given input pptx file for images. Each image will be send to google cloud 
vision API for reverse image search. The results will be added to a new output pptx file, where each slide references a 
slide of the input pptx, listing all contained images and their most likely sources as returned by google cloud vision.


## Setup
Install dependencies:

```shell
$ pip install -r requirements.txt
```

Follow this [tutorial](https://cloud.google.com/vision/docs/libraries?hl=de#client-libraries-install-python) to create a google cloud vision key file for authentication.


## Execution
Call the following python script with your key json file and input/output-pptx files.
```shell
$ python src/main.py -i /path/to/your/input.pptx -o /path/to/your/output.pptx -k /path/to/your/cloud_vision_keyfile.json
```



## License

[MIT License](LICENSE)
