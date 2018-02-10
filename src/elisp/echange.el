;;; echange.el --- MS Exchange emails and calendar events helpers

;; Copyright (C) 2018 Kirill Gerasimenko

;; Author: Kirill Gerasimenko <kirill.gerasimenko@internet-mail.org>
;; Created: 10 Feb 2018
;; Keywords: exchange ews outlook email messages mail calendar
;; URL: https://github.com/kirill-gerasimenko/echange
;; Version: 0.1.0
;; Package-Requires: ((emacs "24.4"))

;; This file is not part of GNU Emacs.

;; This program is free software: you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:

;; Emacs package that provides abilities to open Microsoft Exchange emails
;; from org-mode captured links and adds events for today's and tomorrow's
;; upcoming events to org-agenda.

;;; Code:

(setq lexical-binding t)

(require 'request)
(require 'deferred)
(require 'request-deferred)
(require 'json)
(require 'dash)

(setq request-backend 'url-retrieve)
(setq request-log-level -1)
(setq request-message-level -1)

(defvar echange-calendar-file nil "Path to save calendar org file")
(defvar echange-exchange-base-url nil "Base URL of Exchange EWS server. Used to form email's full URL if it is requested to open email message in browser instead of Outlook")
(defvar echange-server-path "./echange.bat" "Http server executable path")
(defvar echange-server-port "5000" "Http server port")
(defvar echange-exchange-dirs ["Inbox" "Archive"] "Exchange folders names to look for messages")

(defvar echange--session-id nil "Session id used to perform calls to echange http server. Obtained from logon method")

(defun echange--url ()
    (format "http://localhost:%s" echange-server-port))

(defun echange--proc-filter (process text on-start)
  (when (string-match "Started server on port" text)
    (funcall on-start process)))

(defun echange--proc-sentinel (process status on-stop)
  (let* ((exit-statuses '("killed" "interrupt" "finished" "exited abnormally with code"))
         (has-exit-status (seq-contains
                           exit-statuses
                           status
                           (lambda (s1 s2)
                             (string-prefix-p s2 s1 t)))))
    (when has-exit-status
      (funcall on-stop process (substring status 0 (- (length status) 1))))))

(defun echange--ensure-http-service ()
  (if (get-process "echange-http-server")
      (deferred:succeed)
    (deferred:$
      (echange--start-http-service)
      (deferred:error it
        (lambda (err)
          (echange--reset-session)
          (deferred:fail err))))))

(defun echange--start-http-service ()
  (lexical-let ((d (deferred:new))
                (p (start-process "echange-http-server" nil echange-server-path echange/server-port)))
    (set-process-sentinel
     p
     (lambda (process status)
       (echange--proc-sentinel
        process
        status
        (lambda (process status)
          (deferred:errorback-post d (format "Process failed with %S" status))))))
    (set-process-filter
     p
     (lambda (process text)
       (echange--proc-filter
        process
        text
        (lambda (process)
          (deferred:callback-post d)))))
    d))

(defun echange--reset-session ()
  (setq echange--session-id nil))

(defun echange--set-session (session-id)
  (when (not session-id)
    (error "session-id must not be nil"))
  (setq echange--session-id session-id))

(defun echange--get-session ()
  echange--session-id)

(defun echange--parse-response (resp-alist)
  (let* ((value (assoc-default 'value resp-alist))
         (err (assoc-default 'error resp-alist)))
    `(:value ,value :error ,err)))

(defun echange--on-success (response)
  (let ((status-code (request-response-status-code response))
        (response-data (request-response-data response)))
    (cond
     ((= status-code 200)
      (-let (((&plist :value value :error error) (echange--parse-response response-data)))
        (if (equal error t)
            (deferred:fail value)
          (deferred:succeed value))))
     ((not (= status-code 200)) (deferred:fail (request-response-error-thrown response))))))

(defun echange--logon (email password)
  (deferred:$
    (request-deferred
     (concat (echange--url) "/logon")
     :type "POST"
     :headers '(("Content-Type" . "application/x-www-form-urlencoded"))
     :data (format "email=%s&password=%s" email password)
     :parser 'json-read)
    (deferred:nextc it 'echange--on-success)))

(defun echange--calendar2days (session-id)
  (deferred:$
    (request-deferred
     (concat (echange--url) (format "/calendar2days?session-id=%s" session-id))
     :type "GET"
     :parser 'json-read)
    (deferred:nextc it 'echange--on-success)))

(defun echange--entry-id (session-id message-id folders)
  (let* ((msg-id (url-encode-url message-id))
         (fs (string-join (mapcar 'url-encode-url folders) ";"))
         (data (format "session-id=%s&message-id=%s&folders=%s" session-id msg-id fs)))
    (deferred:$
      (request-deferred
       (concat (echange--url) "/entry-id")
       :type "POST"
       :headers '(("Content-Type" . "application/x-www-form-urlencoded"))
       :data data
       :parser 'json-read)
      (deferred:nextc it 'echange--on-success))))

(defun echange--message-url (session-id message-id folders)
  (let* ((msg-id (url-encode-url message-id))
         (fs (string-join (mapcar 'url-encode-url folders) ";"))
         (data (format "session-id=%s&message-id=%s&folders=%s" session-id msg-id fs)))
    (deferred:$
      (request-deferred
       (concat (echange--url) "/message-url")
       :type "POST"
       :headers '(("Content-Type" . "application/x-www-form-urlencoded"))
       :data data
       :parser 'json-read)
      (deferred:nextc it 'echange--on-success))))

;;;###autoload
(defun echange-logon ()
  (interactive)
  (let ((sid (echange--get-session)))
    (if (not sid)
        (condition-case e
            (progn
              (echange--set-session :in-progress)
              (lexical-let ((email (read-string "(ECHANGE) Enter exchange email: "))
                            (password (password-read "(ECHANGE) Enter exchange password: " nil)))
                (deferred:$
                  (echange--ensure-http-service)
                  (deferred:nextc it
                    (lambda ()
                      (echange--logon email password)))
                  (deferred:nextc it
                    (lambda (session-id)
                      (echange--set-session session-id)
                      (deferred:succeed session-id)))
                  (deferred:error it ; if error - set session to nil and forward error
                    (lambda (err)
                      (echange--reset-session)
                      (deferred:fail err))))))
          (quit
           (progn
             (echange--reset-session)
             (deferred:fail "Cancelled action")))
          (error
           (echange--reset-session)
           (deferred:fail e)))
      (if (equal sid :in-progress)
          (deferred:fail "Logon is already in progress")
        (deferred:succeed sid)))))

;;;###autoload
(defun echange-open-message (message-id &optional in-browser)
  (interactive)
  (lexical-let ((message-id message-id) ; for some reason variables didn't get closed over in lambda below
                (in-browser in-browser)); lexical-let is the only way I've found this to work for now
    (deferred:$
      (echange-logon)
      (deferred:nextc it
        (lambda (session-id)
          (if (not in-browser)
              (deferred:$
                (echange--entry-id session-id message-id echange-exchange-dirs)
                (deferred:nextc it
                  (lambda (entry-id)
                    (w32-shell-execute "open" "outlook" (concat "outlook:" entry-id)))))
            (deferred:$
              (echange--message-url session-id message-id echange-exchange-dirs)
              (deferred:nextc it
                (lambda (message-url)
                  (let ((full-url (concat echange-exchange-base-url (url-encode-url message-url))))
                    (w32-shell-execute "open" full-url))))))))
      (deferred:error it
        (-lambda ((&plist 'error error))
          (message "(ECHANGE) Error getting email entry-id: %S" error))))))


;;;###autoload
(defun echange-calendar2days ()
  (interactive)
  (deferred:$
    (echange-logon)
    (deferred:nextc it
      (lambda (session-id)
        (echange--calendar2days session-id)))
    (deferred:nextc it
      (lambda (content)
        (with-temp-buffer
          (insert "#+CATEGORY: work")
          (newline)
          (insert "* Calendar")
          (newline)
          (insert content)
          (write-file echange-calendar-file))))
    (deferred:error it
      (-lambda ((&plist 'error error))
        (message "(ECHANGE) Error getting calendar for 2 days: %S" error)))))

;;;###autoload
(defun echange-logoff ()
  (interactive)
  (lexical-let ((sid (echange--get-session)))
    (when (and sid (not (equal sid :in-progress)))
      (echange--reset-session)
      (deferred:$
        (echange--ensure-http-service)
        (deferred:nextc it
          (lambda ()
            (request-deferred
             (concat (echange--url) "/logoff")
             :type "POST"
             :headers '(("Content-Type" . "application/x-www-form-urlencoded"))
             :data (format "session-id=%s" sid)
             :parser 'json-read)))
        (deferred:error it
          (-lambda ((&plist 'error error))
            (message "(ECHANGE) Error logging off: %S" error)))))))

(provide 'echange)

;;; echange.el ends here
