# From-etutcl-to-calendar
Just an app to convert data from [etutcl](http://etutcl.fr) and import it into a google calendar

### Requirements
Docker client installed

### How to run 
+ Create a ~/.etutcl/config/config.json file and setup your configuration here according to the config.json.example file
+ In ~/.etutcl/data/credentials/ put your client_secret.json given by google's OAUTH API
+ All the app's data should be stored in ~/.etutcl/ (credentials and config)
+ Finally run `docker run --name etutcl -v ~/.etutcl:/data r.bde-insa-lyon.fr/al26p/from-etutcl-to-calendar`

### Licence
Ce projet est sous GNU GPL-3.0

© 2018 Alban PRATS