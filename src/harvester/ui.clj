(ns harvester.ui
  (:require
   [harvester.analysis :as an]
   [net.cgrand.enlive-html :as html])
  (:use seesaw.core
        seesaw.chooser
        seesaw.swingx
        seesaw.table
        seesaw.util
        seesaw.font)
  (:import [java.io.Reader]
           (org.apache.poi.hssf.usermodel HSSFPalette HSSFWorkbook)
           (org.apache.poi.hssf.record PaletteRecord)))

(def status (label "Ready"))

(def html-text (text :text "<html><body><input id=\"hi\" name=\"name\" /></body></html>"
                     :editable? true))
(def url-text (text "https://pnc.starssmp.com/InitialRegistration.aspx"))

(def exit-action (action :handler dispose! :name "Exit"))

(defn- url-table []
  (horizontal-panel
    :border [5 "Request Result"]
    :id :table_panel
    :items [(scrollable (table-x :id :table
                                 :model @an/table-modl)) ]))

(defn save-as-act [c f]
  (an/write-inputs-to-excel (an/table-to-xls-model) f))

(def text-action
  (border-panel
    :border 5
    :north (toolbar :items [exit-action])
    :center
    (border-panel
      :north
      (horizontal-panel
        :border [5 "Configure Request"]
        :items [(action :handler
                        (fn [s]
                          (choose-file :type :open
                                       :success-fn (fn [c f] (reset! an/file-data (an/into-xml (html/html-resource f)))))
                          (an/harvester-file-to-table)
                          (config! (select text-action [:#table]) :model @an/table-modl))
                        :name "Load File")
                (action :handler (fn [e] (choose-file :type :save :success-fn save-as-act))
                        :name "Save As...")])
      :center
      (url-table))
    :south status))

(defn h-action [title desc content]
  (border-panel
    :id :h-action
    :hgap 5 :vgap 5 :border 5
    :north (label (str title " - " desc))
    :center content))

(def h-actions
 (delay
    {:plain-text (h-action "Plain HTML" "Paste plain HTML in the text box to process."
                           text-action)}))

(defn main-panel []
  (frame :title "Harvester"
         :size [800 :by 600]
         :content
         (border-panel
           :hgap 5 :vgap 5 :border 5
           :north (label :text "Please select how to process the HTML.")
           :center
           (left-right-split
             (listbox
               :id :htmltype
               :model (keys @h-actions))
             (border-panel
               :id :actioncontainer
               :center (:plain-text @h-actions))))))

(defn behaviors [root]
  (let [html-type (select root [:#htmltype])
        container (select root [:#actioncontainer])]
    (listen html-type :selection
            (fn [e]
              (replace! container
                        (select container [:#h-action])
                        (@h-actions (selection html-type))))))
  root)
