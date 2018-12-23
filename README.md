# forecast_dataset_etl
A etl project to pre-handler the dataset for Amazon Forecast

## Getting Started
---
- Prepare the dataset in excel version over 2007, which suffix is **xlsx**
- Sort the dataset in order by product name
- Any modification to the specific field if needed

## Execution
---
```
python main.py --import-file sample.xlsx --output-folder output
```
- import-file: the file which is going to be transformed
- output-folder: the output folder to store all the transformed files