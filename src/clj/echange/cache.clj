(ns echange.cache
  (:require [failjure.core :as f]))

(defn create [] (atom {}))

(defn put
  [cache key value]
  (swap! cache assoc key value))

(defn value
  ([cache key] (get @cache key))
  ([cache key value-factory]
   (let [val (get @cache key)]
     (if (some? val)
       val
       (f/when-let-ok? [new-val (f/try* (value-factory))]
         (swap! cache assoc key new-val)
         new-val)))))

(defn reset
  ([cache] (reset! cache {}))
  ([cache key] (swap! cache dissoc key)))


