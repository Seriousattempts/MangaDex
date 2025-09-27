# Exporting your MangaDex library

## Prerequisites

Your MangaDex
- Username
- Password
- personal client id
- Secret

You get your personal client id and secret under API Clients in Settings https://mangadex.org/settings by creating them https://api.mangadex.org/docs/02-authentication/personal-clients/

Have python with the following installed:
- pandas
- requests


# 1. Run *Export libraries from MangaDex.py*

- Paste your MangaDex information at the top
- Choose a location to export the .xlsx files to
- Choose what you want to export. You can export individual files or your entire library


# 2. Choose the external website to export your files to:
MAL: Grab your user_id by exporting your list (https://myanimelist.net/panel.php?go=export), extracting the .gz file and opening the .xml file with notes or another program. Format should look like this:

```
		<myinfo>
			<user_id>123456</user_id>
			<user_name>YourMALname</user_name>
			<user_export_type>2</user_export_type>
			<user_total_manga>2345</user_total_manga>
			<user_total_reading>678</user_total_reading>
			<user_total_completed>901</user_total_completed>
			<user_total_onhold>234</user_total_onhold>
			<user_total_dropped>567</user_total_dropped>
			<user_total_plantoread>0</user_total_plantoread>
		</myinfo>
```

Run *Convert .xlsx to .xml MAL.py* and use that user_id. Choose the proper .xlsx file with the proper status and convert to .xml.

The following sites use .xml to import:

- AniList: https://anilist.co/settings/import
- Anime-Planet: https://www.anime-planet.com/users/import_list.php
- Kitsu: https://kitsu.app/settings/imports
- MAL: https://myanimelist.net/import.php

For Mangaupdates, you run *Mangaupdates MD List.py* and use your normal login details.


# Don't forget to support the Author/Artist


## YouTube Visual version:

[![YouTube Visual version](https://i.ytimg.com/vi/u0VuEufNFfY/maxresdefault.jpg)](https://www.youtube.com/watch?v=u0VuEufNFfY)

# - B O N U S -

Anime recommendations:
- https://www.reddit.com/user/theanimesyndicate/m/completedanimecollection_can/
- https://discord.gg/RQsUfqNXJ9

Manga Recommendations: 
- https://www.reddit.com/user/mangasyndicate/m/completedmangacollection_can/
- https://discord.gg/SMzR6zzXFu
