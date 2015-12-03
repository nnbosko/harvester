(ns harvester.core
  (:require
    [clojure.zip :as zip]
    [clojure.pprint :as pp]
    [clojure.string :refer [trim]]
    [zip.visit :refer :all]
    [net.cgrand.enlive-html :as html]
    [dk.ative.docjure.spreadsheet :as sheet])
  (:use seesaw.core
        seesaw.chooser
        seesaw.swingx
        seesaw.table
        seesaw.util
        seesaw.font
        [clojure.java.io :only [file output-stream reader writer]])
  (:import [java.io.Reader]
           (org.apache.poi.hssf.usermodel HSSFPalette HSSFWorkbook)
           (org.apache.poi.hssf.record PaletteRecord))
  (:gen-class))

(def file-data (atom ""))
(def site-map (atom {}))
(def site-zipper (atom {}))

(defn into-xml
  [coll]
  (if-not (string? coll)
    (let [tag (into {} coll)]
      (update-in tag [:content] (partial map into-xml)))
    coll))

(defn get-html-from-site [site]
  (let [res (-> site java.net.URL. html/html-resource)]
    (into-xml res)))

(defn load-test []
  (reset! site-map (get-html-from-site "https://pnc.starssmp.com/InitialRegistration.aspx"))
  (reset! site-zipper (zip/xml-zip @site-map)))

(defn load-zipper [url]
  (reset! site-map (get-html-from-site url))
  (reset! site-zipper (zip/xml-zip @site-map)))

(defvisitor collect-inputs :pre [n s]
            (if (= :input (:tag n)) {:state (let [attrs (:attrs n)
                                                  name (:name attrs)
                                                  id (:id attrs)
                                                  typ (:type attrs)]
                                              (if (not= typ "hidden")
                                                (conj s {:node n})))}))

(defvisitor collect-inputs-wb :pre [n s]
            (if (and (= :input (:tag n))
                     (not= "hidden" (:type (:attrs n))))
              {:state (let [attrs (:attrs n)
                            name (:name attrs)
                            id (:id attrs)
                            typ (:type attrs)]
                            (conj s [typ id name]))}))

(defvisitor check-nearby-text :pre [n s]
            (let [pn (zip/up n)]
              (if (or (= :span (:tag pn))
                      (= :label (:tag pn))
                      (= :font (:tag pn)))
                (assoc n :possible-text? s))))

(defvisitor collect-selects :pre [n s]
            (if (= :select (:tag n)) {:state (conj s n)}))

(defvisitor collect-select-attrs-wb :pre [n s]
            (if (= :select (:tag n))
              {:state (let [attrs (:attrs n)
                            name (:name attrs)
                            id (:id attrs)]
                        (conj s ["select" id name]))}))

(defvisitor collect-inputs-selects-wb :pre [n s]
            (if (or
                  (= :select (:tag n))
                  (= :input (:tag n)))
              (case (:tag n)
                :select {:state (let [attrs (:attrs n)
                                      name (:name attrs)
                                      id (:id attrs)]
                                  (conj s ["select" id name]))}
                :input {:state (let [attrs (:attrs n)
                                     name (:name attrs)
                                     id (:id attrs)
                                     typ (:type attrs)]
                                 (conj s [typ id name]))})))


(defvisitor collect-inputs-selects-table :pre [n s]
            (if (or
                  (= :select (:tag n))
                  (= :input (:tag n))
                  (= :textarea (:tag n))
                  (= :option (:tag n))
                  )
              (case (:tag n)
                :select {:state (let [attrs (:attrs n)
                                      name (:name attrs)
                                      id (:id attrs)
                                      opts (:content n)]
                                  (conj s {:type "select" :id id :name name
                                           :code (if (not= nil id)
                                                   (str "dd_select_n('" id "', driver, data[''])")
                                                   (str "dd_select_n('" name "', driver, data[''], 'name')"))}))}
                :option {:state (let [optval (-> n :attrs :value str)
                                      optval-first (-> n :attrs :value first str)
                                      opttxt (-> n :content first str)
                                      is-val-string? (= java.lang.String (type optval))
                                      optval-output (if is-val-string? optval optval-first)]
                                  ;(println "optval: " optval ", type: " (type optval))
                                  (conj s {:type "-- option"
                                           :id optval-output
                                           :name opttxt
                                           :code (str " { \"val\" : \"" optval-output "\",  \" txt \" : \"" opttxt "\"}") }))}
                :input {:state (let [attrs (:attrs n)
                                     name (:name attrs)
                                     id (:id attrs)
                                     typ (:type attrs)]
                                 (conj s {:type typ :id id :name name
                                          :code (case typ
                                                  "text" (if (not= nil id)
                                                           (str "t_type_n('" id "', driver, data[''])")
                                                           (str "t_type_n('" name "', driver, data[''], 'name')"))
                                                  "radio" (if (not= nil id)
                                                            (str "click_n('" id "', driver)")
                                                            (str "click_n('" name "', driver, 'name')"))
                                                  "checkbox" (if (not= nil id)
                                                            (str "click_n('" id "', driver)")
                                                            (str "click_n('" name "', driver, 'name')"))
                                                  "submit" (if (not= nil id)
                                                               (str "click_n('" id "', driver)")
                                                               (str "click_n('" name "', driver, 'name')"))
                                                  "button" (if (not= nil id)
                                                             (str "click_n('" id "', driver)")
                                                             (str "click_n('" name "', driver, 'name')"))
                                                  "hidden" "")
                                          }))}
                :textarea {:state (let [attrs (:attrs n)
                                     name (:name attrs)
                                     id (:id attrs)]
                                   (conj s {:type "textarea" :id id :name name
                                            :code (if (not= nil id)
                                                    (str "t_type_n('" id "', driver, data[''])")
                                                    (str "t_type_n('" name "', driver, data[''], 'name')"))}))})))

(defn load-site [site]
  (reset! site-map (get-html-from-site site)))

(defn write-map-to-file [map file]
  (pp/pprint
    map (clojure.java.io/writer (str "/home/nicolas/Work/CloudCustom/harvester/outputs/" file))))

(defn set-custom-colors [wb]
  (let [sel-col (sheet/color-index :bright_green)
        opt-col (sheet/color-index :green)
        wb-p (.getCustomPalette wb)
        ]
    #_(.set!)
    (.setColorAtIndex wb-p sel-col 180 167 214)
    (.setColorAtIndex wb-p opt-col 217 210 233)))

(defn write-inputs-to-excel [dat file]
  (let [wb (sheet/create-xls-workbook "Inputs" dat)
        sheet (sheet/select-sheet "Inputs" wb)
        row-sequence (sheet/row-seq sheet)]
    (set-custom-colors wb)
    (.setColumnWidth sheet 0 2000)
    (.setColumnWidth sheet 1 7500)
    (.setColumnWidth sheet 2 7500)
    (.setColumnWidth sheet 3 20000)
    (doseq [r row-sequence]
      (case (str (first r))
        "Type" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :yellow, :font {:bold true}}))
        "text" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :pale_blue}))
        "radio" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :light_cornflower_blue}))
        "checkbox" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :light_cornflower_blue}))
        "submit" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :light_cornflower_blue}))
        "button" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :light_cornflower_blue}))
        "hidden" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :grey_40_percent}))
        "textarea" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :pale_blue}))
        "select" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :bright_green}))
        "-- option" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :green})))
      )
    #_(doseq
      [item row-sequence]
      (case (:type item)))
    (sheet/save-workbook! (str file) wb)
    ))

  #_(let [dat (:state (visit @site-zipper [] [collect-inputs-selects-wb]))
        wb (sheet/create-workbook "Inputs" dat)
        sht (sheet/select-sheet "Inputs" wb)
        hdr (first (sheet/row-seq sht))]
    (do
      (sheet/set-row-style! hdr (sheet/create-cell-style! wb {:background :yellow, :font {:bold true}}))
      (sheet/save-workbook! (str "/home/nicolas/Work/CloudCustom/harvester/outputs/tests.xlsx") wb)))

(def table-rows (atom {}))
(def table-modl (atom (table-model :columns [{:key :type :text "Type"}
                                             {:key :id :text "ID"}
                                             {:key :name :text "Name"}
                                             {:key :code :text "LOC"}]
                                   :rows [{:type "Type 1"
                                           :id "ID 1"
                                           :name "Name 1"
                                           :code "iwait(0.1)"}])))

(defn table-to-xls-model []
  (let [dat (value-at @table-modl (range 0 (row-count @table-modl)))
        rv (into [["Type" "ID" "Name" "LOC"]] (mapv (fn [e] [(:type e) (:id e) (:name e) (:code e)]) dat))]
    rv
    ))

(defn reload-table-model [zipper]
  (let [row-res (:state (visit @zipper [] [collect-inputs-selects-table]))
        mod-res (table-model :columns [{:key :type :text "Type"}
                                       {:key :id :text "ID"}
                                       {:key :name :text "Name"}
                                       {:key :code :text "LOC"}]
                             :rows row-res)]
    (reset! table-rows row-res)
    (reset! table-modl mod-res)))

(defn harvester-file-to-table []
  (reset! site-zipper (zip/xml-zip @file-data))
  (reload-table-model site-zipper))

(defn harvester-url-to-table [url]
  (load-zipper (str url))
  (reload-table-model site-zipper))

(defn do-site-request [url-str f]
  (future
    (let [result (if-let [url (to-url url-str)]
                   (harvester-url-to-table url)
                   "Invalid URL")]
      (invoke-later (f result)))))

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
                                 :model @table-modl)) ]))

(defn save-as-act [c f]
  (write-inputs-to-excel (table-to-xls-model) f))

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
                                       :success-fn (fn [c f] (reset! file-data (into-xml (html/html-resource f)))))
                          (harvester-file-to-table)
                          (config! (select text-action [:#table]) :model @table-modl))
                        :name "Load File")
                (action :handler (fn [e] (choose-file :type :save :success-fn save-as-act))
                        :name "Save As...")])
      :center
      (url-table))
    :south status))


(def url-action
  (border-panel
    :border 5
    :north (toolbar :items [exit-action])
    :center
    (border-panel
      :north
      (horizontal-panel
        :border [5]
        :items ["URL" url-text
                (action :handler (fn [s]
                                   (text! status "Busy")
                                   (do-site-request (text url-text)
                                               (fn [s]
                                                 (text! status "Ready")
                                                 (config! (select url-action [:#table]) :model @table-modl)))) :name "Go")
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
                           text-action)
     :url (h-action "URL" "Put a URL here to process."
                     url-action)}))

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

(defn -main [& args]
  (-> (main-panel)
      behaviors
      pack!
      show!))