FROM rocker/rstudio
USER root
ENV PATH "/usr/lib/rstudio-server/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"

# Attempt to fix potential broken mismatches (intermittent failures)
RUN apt-get -y clean
RUN rm -Rf /var/lib/apt/lists
RUN echo "Acquire::By-Hash \"yes\"; ">/etc/apt/apt.conf.d/01byhash
RUN echo "Acquire::http::Pipeline-Depth 0;" > /etc/apt/apt.conf.d/99fixbadproxy
RUN echo "Acquire::http::No-Cache true;" >> /etc/apt/apt.conf.d/99fixbadproxy
RUN echo "Acquire::BrokenProxy    true;" >> /etc/apt/apt.conf.d/99fixbadproxy
RUN echo "deb http://debian.csail.mit.edu/debian/ stable main" > /etc/apt/sources.list

# Install stuff we can through apt-get
RUN apt-get -y update
RUN apt-get -y install default-jre-headless default-jdk-headless r-cran-shiny r-cran-markdown libxml2-dev texlive xzdec

# Allow texmf (markup) to write to this directory in RStudio
RUN chmod 777 /var/lib/texmf
RUN echo "install.packages(c('shinyjs', 'shinyalert','readxl','rmarkdown','dplyr', 'mailR', 'futile.logger', 'properties', 'kableExtra', 'caTools', 'bitops', 'rprojroot', 'png'))" | R -

# Cleanup downloaded packages to reclaim some space so this doesn't take so much
RUN apt-get -y clean