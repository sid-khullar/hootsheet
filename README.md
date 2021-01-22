These are scattered notes and stuff. If you think you can use any of this and need help figuring it out, email me at siddhartha.khullar at gee mail.
# hootsheet
Collates messages based on specified sheet, number of messages and starting date into separate sheets. Written to simplify the tedious job of creating multiple bulk sending CSV files for Hootsuite, especially if one has different combinations of target social networks for different message sets, and the limit of 350 messages is too small.

The script needs a config sheet where to read the message configuration, requires the entries under the 'Sheets' column to exist with the standard Hootsuite CSV format, and needs sheets named in the Tag column to exist.

The idea is to have a list of different content streams as source sheets, which can be added to over time, and writes output to the specified (tag) sheets, which can be exported as CSV and imported into the Hootsuite bulk message uploader. Tags are separated based on the target networks.

Sheet	Messages	Networks	Total Messages	Begin Date	Tag	Targets

# Ad Rotate
Used with the Ad Inserter plugin for Wordpress. I use it to insert an ad after the first para of a blog post and wanted to randomise the image displayed.
