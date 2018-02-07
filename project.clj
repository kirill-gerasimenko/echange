(defproject echange "0.1.0-SNAPSHOT"
  :description "Microsoft EWS accessing layer for echange emacs plugin"
  :url "https://github.com/kirill-gerasimenko/echange"
  :source-paths ["src/clj"]
  :plugins [[lein-ring "0.12.3"]]
  :ring {:handler echange.app/main
         :nrepl {:start? true}}
  :dependencies [[com.microsoft.ews-java-api/ews-java-api "2.0"]
                 [org.clojure/clojure "1.9.0"]
                 [failjure "1.2.0"]
                 [compojure "1.6.0"]
                 [ring/ring-core "1.6.3"]
                 [ring/ring-jetty-adapter "1.6.3"]
                 [ring/ring-devel "1.6.3"]
                 [org.clojure/data.json "0.2.6"]
                 [clojure.java-time "0.3.1"]])
