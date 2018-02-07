(ns echange.core
  (:import microsoft.exchange.webservices.data.core.service.item.EmailMessage
           java.text.SimpleDateFormat)
  (:require [echange.exchange :as exch]
            [clojure.pprint :as p]
            [failjure.core :as f]
            [echange.cache :as cache]
            [java-time :as t]
            [echange.utils :as u]
            [clojure.string :as s]))

(def ^:private cache (cache/create))

(defn ^:private service
  [email password]
  (let [create-service #(f/when-let-ok? [s (exch/service email password)]
                          {:password password :service s})]
    (f/when-let-ok? [cached (cache/value cache email create-service)]
      (if (= password (:password cached))
        cached
        (f/fail "Password is incorrect")))))

(defn ^:private find-folders
  [service folders]
  (cache/value cache folders #(exch/find-folders service folders)))

(defn session->service
  [session-id]
  (if-some [service (some->> session-id
                            (cache/value cache)
                            :email
                            (cache/value cache)
                            :service)]
    service
    (f/fail "Session with id %s is not found" session-id)))

(defn logon [email password]
  (f/when-let-ok? [_ (service email password)]
    (let [session-id (u/uuid)]
      (cache/put cache session-id {:email email})
      session-id)))

(defn logoff [session-id]
  (cache/reset cache session-id)
  nil)

(defn message-id->entry-id
  [session-id message-id folders]
  (f/attempt-all
      [service (session->service session-id)
       email (->> session-id (cache/value cache) :email)
       folders (find-folders service folders)
       message (exch/find-message service message-id folders)
       entry-id (exch/entry-id service email (.. message getId getUniqueId))]
      entry-id))

(defn message-id->url
  [session-id message-id folders]
  (f/attempt-all
      [service (session->service session-id)
       email (->> session-id (cache/value cache) :email)
       folders (find-folders service folders)
       message (exch/find-message service message-id folders)
       url (exch/relative-url service message)]
    url))

(defn calendar-2-days
  [session-id]
  (f/attempt-all
      [service (session->service session-id)
       now (t/local-date-time)
       start (t/truncate-to now :days)
       end (-> now
               (t/plus (t/days 2))
               (t/truncate-to :days))
       events (exch/calendar-in-range service start end)]
    (->> events
         (filter #(t/after? (:to %) now))
         (map (fn [{:keys [from to subject sender]}]
                (let [fmt "yyyy-MM-dd EEE HH:mm"
                      scheduled (format "SCHEDULED: <%s>" (t/format fmt from))
                      deadline (format "DEADLINE: <%s>" (t/format fmt to))
                      dates (if (t/after? from now)
                              (str scheduled " " deadline)
                              deadline)]
                  (format "** %s (%s)\n   %s" subject sender dates))))
         (s/join "\n"))))
