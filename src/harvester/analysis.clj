(ns harvester.analysis
  (:require
    [clojure.zip :as zip]
    [clojure.pprint :as pp]
    [clojure.string :refer [trim]]
    [zip.visit :refer :all]
    [net.cgrand.enlive-html :as html]
    [dk.ative.docjure.spreadsheet :as sheet])
  (:use seesaw.table)
  (:import [java.io.Reader]
           (org.apache.poi.hssf.usermodel HSSFPalette HSSFWorkbook)
           (org.apache.poi.hssf.record PaletteRecord)))

(def file-data (atom ""))
(def site-map (atom {}))
(def site-zipper (atom {}))
(def table-rows (atom {}))
(def table-modl (atom (table-model :columns [{:key :type :text "Type"}
                                             {:key :id :text "ID"}
                                             {:key :name :text "Name"}
                                             {:key :code :text "LOC"}]
                                   :rows [{:type "Type 1"
                                           :id "ID 1"
                                           :name "Name 1"
                                           :code "iwait(0.1)"}])))


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
        wb-p (.getCustomPalette wb)]

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
        "-- option" (sheet/set-row-style! r (sheet/create-cell-style! wb {:background :green}))))
    (sheet/save-workbook! (str file) wb)))

(defn table-to-xls-model []
  (let [dat (value-at @table-modl (range 0 (row-count @table-modl)))
        rv (into [["Type" "ID" "Name" "LOC"]] (mapv (fn [e] [(:type e) (:id e) (:name e) (:code e)]) dat))]
    rv))

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
