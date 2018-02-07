(ns echange.exchange
  (:require [failjure.core :as f]
            [echange.utils :as u]
            [java-time :as t])
  (:import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl
           [microsoft.exchange.webservices.data.core ExchangeService PropertySet]
           [microsoft.exchange.webservices.data.core.service.schema EmailMessageSchema FolderSchema]
           [microsoft.exchange.webservices.data.core.enumeration.property BasePropertySet WellKnownFolderName]
           [microsoft.exchange.webservices.data.search FolderView ItemView CalendarView]
           microsoft.exchange.webservices.data.core.enumeration.misc.IdFormat
           microsoft.exchange.webservices.data.core.enumeration.search.LogicalOperator
           microsoft.exchange.webservices.data.credential.WebCredentials
           microsoft.exchange.webservices.data.misc.id.AlternateId
           microsoft.exchange.webservices.data.property.definition.PropertyDefinitionBase
           microsoft.exchange.webservices.data.property.definition.StringPropertyDefinition
           [microsoft.exchange.webservices.data.search.filter SearchFilter$IsEqualTo SearchFilter$SearchFilterCollection]))

(def ^:private true-validator
  (reify IAutodiscoverRedirectionUrl
    (autodiscoverRedirectionUrlValidationCallback [_ _] true)))

(defn message-id-filter
  [messageid]
  (SearchFilter$IsEqualTo. EmailMessageSchema/InternetMessageId messageid))

(defn folder-filter
  [folder-name]
  (SearchFilter$IsEqualTo. FolderSchema/DisplayName folder-name))

(defn combine-filters
  [operation filters]
  (SearchFilter$SearchFilterCollection. operation filters))

(defn service
  [email password]
  (f/try*
   (doto (ExchangeService.)
     (.setUseDefaultCredentials false)
     (.setCredentials (WebCredentials. email password))
     (.autodiscoverUrl email true-validator))))

(defn find-folders
  [service folders]
  (let [folder-count (count folders)]
    (if (> folder-count 0)
       (let [view (FolderView. folder-count)
             filter (->> folders
                         (map folder-filter)
                         (combine-filters LogicalOperator/Or))]
         (f/try*
           (let [folders-result (.findFolders service WellKnownFolderName/MsgFolderRoot filter view)]
             (if (= 0 (count (.getFolders folders-result)))
               (f/fail "Exchange folders are not found")
               folders-result))))
       (f/fail "No folders provided for search"))))

(defn find-message
  [service message-id folders]
  (let [item-view (ItemView. 1)
        message-filter (message-id-filter message-id)]
    (f/try*
     (if-some [message
               (->> folders
                    (map #(.findItems service (.getId %) message-filter item-view))
                    (mapcat #(.getItems %))
                    (first))]
       message
       (f/fail "Message '%s' is not found" message-id)))))

(defn entry-id
  [service mailbox ews-id]
  (let [a-id (AlternateId. IdFormat/EwsId ews-id mailbox false)]
    (f/try*
     (-> (.convertId service a-id IdFormat/HexEntryId)
         (.getUniqueId)))))

(defn calendar-in-range
  [service start end]
  (f/try*
   (->> (CalendarView. (u/local->java-date start) (u/local->java-date end))
        (.findAppointments service WellKnownFolderName/Calendar)
        (.getItems)
        (map (fn [i]
               {:subject (.getSubject i)
                :sender (.. i getOrganizer getName)
                :from (-> i .getStart u/java-date->local)
                :to (-> i .getEnd u/java-date->local)
                :duration (.getDuration i)})))))

(defn relative-url
  "Gets relative URL of the email messsage for OWA access"
  [service message]
  (let [property-set (doto (PropertySet.)
                       (.setBasePropertySet BasePropertySet/IdOnly)
                       (.add EmailMessageSchema/WebClientReadFormQueryString))]
    (f/try* (.load message property-set)
            (.getWebClientReadFormQueryString message))))
