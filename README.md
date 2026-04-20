# gt-excel-writer

A write-only Excel datastore for GeoTools.


## development

Run tests using `mvn test -Dlogging-profile=verbose-logging`

Run QA procedures using `mvn -B -fae clean install -Dspotless.action=check -Dpom.fmt.action=verify -Dqa=true -DskipTests=true`
