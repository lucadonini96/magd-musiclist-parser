# magd-musiclist-parser
Python script to parse a chapel service music list Word document and produce an Excel table to upload to phpMyAdmin. As used in Magdalene College, Cambridge, UK.

Install the required Python packages with `pip install requirements.txt`

Run `jupyter notebook` in the root directory of this repository.

In the Jupyter notebook, change `inputfile` to the path of the input Word document. Make sure you "accept all changes" in the Word Revision tab before running the Jupyter notebook.

Run all cells in the notebook. You may set the ID of the first service of the list in the command `ml = musiclist(inputfile,firstid=None)` by setting `firstid` to the desired numerical value.

The output Excel table is by default saved to the `output` folder.
