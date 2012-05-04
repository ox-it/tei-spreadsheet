#!/bin/bash

MEDIA_TYPE=`file -b --mime-type "$1"`

case "$MEDIA_TYPE" in
  "application/vnd.ms-excel")
    TEMP_DIR=`mktemp -d`
    unzip "$1" -d $TEMP_DIR > /dev/null
    saxon office-open-xml/main.xsl -it:main url=${TEMP_DIR}/
    rm -rf $TEMP_DIR
    ;;
  *)
    echo "Could not determine file type"
    exit 1
    ;;
esac

