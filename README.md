# WordWebRephraseAddIn
# MS Word Add-In for Rephrasing Using a Fine-Tuned GPT-3 Model.

This Add-In for Microsoft Word is written to use a fine-tuned GPT-3 model for rephrasing a selected sentence.  It returns the top two results.  Initially, the Curie engine was used, but this is more expensive.  The Ada engine produced inferior results, but the Babbage engine worked quite well.  

*Please see the Jupyter Notebook for some guidance on training your own model.  Please note, it is just for reference so you can see the commands. Only the final data for fine-tuning is included.*

The data file for fine-tuning the Curie model is called: output_prepared_10_6_21_filtered.jsonl.

The final data file used for fine-tuning the Babbage engine, which was the most affordable and functional model is: output_10_16_21_filtered.jsonl.

# Installation

You need to obtain your OpenAI API Key and a Fine-Tuned Model ID after training a model.  See the [Jupyter Notebook](Training%20Sentence%20Rephrasing-GitHub.ipynb) for the history of how I did it.  You can skip to the end to train a model on the Babbage engine, which was the best in terms of price of operation and model performance.

*Edit Home.js to add your OpenAI API Key and Fine-Tuned Model ID*

# Screenshot

![Screenshot of the Add-In running in MS Word](rephrase_screenshot.png?raw=true "Screenshot")
