FROM rocker/rstudio
USER root
ENV PATH "/usr/lib/rstudio-server/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"

RUN apt-get -y update
RUN apt-get -y install default-jre-headless
RUN apt-get -y install default-jdk-headless
RUN apt-get -y install r-cran-shiny
RUN apt-get -y install r-cran-markdown

RUN echo "install.packages(c('shinyjs', 'shinyalert','readxl','rmarkdown','dplyr', 'mailR', 'futile.logger', 'properties'))" | R -