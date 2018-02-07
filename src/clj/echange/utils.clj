(ns echange.utils
  (:require [java-time :as t]))

(defn uuid [] (str (java.util.UUID/randomUUID)))

(defn local->java-date
  [local-date-time]
  (-> local-date-time
      (t/zoned-date-time (t/zone-id))
      (t/java-date)))

(defn java-date->local
  [date-time]
  (-> date-time
      (t/zoned-date-time (t/zone-id))
      (t/local-date-time)))
