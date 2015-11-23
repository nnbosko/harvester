(defproject harvester "1.01"
  :description "HTML input/select finder."
  :url "http://github.com/nnbosko"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.7.0"]
                 [org.clojure/data.zip "0.1.1"]
                 [dk.ative/docjure "1.9.0"]
                 [enlive "1.1.6"]
                 [seesaw "1.4.5"]
                 [zip-visit "1.1.0"]]
  :aot [harvester.core]
  :main harvester.core)
