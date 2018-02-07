(ns echange.app
  (:require [compojure.core :refer :all]
            [compojure.route :as route]
            [echange.core :as c]
            [ring.util.response :as r]
            [ring.middleware.params :refer [wrap-params]]
            [clojure.string :as str]
            [failjure.core :as f]
            [clojure.data.json :as json]))

(defn ^:private result->map [result]
  (let [failed (f/failed? result)
        value (if failed (f/message result) result)
        payload {:value value}]
    (if failed
      (assoc payload :error true)
      payload)))

(defn entry-id [session-id message-id folders]
  (let [folders-list (str/split folders #"\;")]
    (-> session-id
        (c/message-id->entry-id message-id folders-list)
        (result->map))))

(defn logon [email password]
  (-> (c/logon email password)
      (result->map)))

(defn logoff [session-id]
  (c/logoff session-id))

(defn calendar2days [session-id]
  (-> (c/calendar-2-days session-id)
      (result->map)))

(defn message-url [session-id message-id folders]
  (let [folders-list (str/split folders #"\;")]
    (-> (c/message-id->url session-id message-id folders-list)
        (result->map))))

(defroutes routes-handler
  (POST "/logon" [email password]
    (-> (logon email password)
        (json/write-str)))
  (POST "/logoff" [session-id]
    (do (logoff session-id)
        (r/response "")))
  (POST "/entry-id" [session-id message-id folders]
    (-> (entry-id session-id message-id folders)
        (json/write-str)))
  (POST "/message-url" [session-id message-id folders]
    (-> (message-url session-id message-id folders)
        (json/write-str)))
  (GET "/calendar2days" [session-id]
    (-> (calendar2days session-id)
        (json/write-str))))

(def main (-> routes-handler
              wrap-params))

