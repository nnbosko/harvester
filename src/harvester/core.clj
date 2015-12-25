(ns harvester.core
  (:require [harvester.ui :as ui])
  (:use seesaw.core)
  (:gen-class))

(defn -main [& args]
  (-> (ui/main-panel)
      ui/behaviors
      pack!
      show!))
