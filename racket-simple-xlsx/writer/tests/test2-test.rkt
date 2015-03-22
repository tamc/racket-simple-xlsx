#lang racket

(require rackunit/text-ui)

(require rackunit "../../main.rkt")

(require "../../lib/lib.rkt")

(define test-test2
  (test-suite
   "test-test2"

   (test-case 
    "test-write-read-number"
    (test-case
     "test-write-read-number"

     (let ([data_list '()]
           [correct_sum 0])
       ;; one thousand lines with random number
       (let loop ([count 1]
                  [rad (round (* (random) 1000))])
         (when (<= count 1000)
               (set! correct_sum (+ correct_sum rad))               
               (set! data_list `(,@data_list (,rad)))
               (loop (add1 count) (round (* (random) 1000)))))

       (set! data_list `((,@data_list)))
       
       (dynamic-wind
           (lambda ()
             (write-xlsx-file data_list #f "test2.xlsx"))
           (lambda ()
             (with-input-from-xlsx-file
              "test2.xlsx"
              (lambda ()
                (load-sheet "Sheet1")

                (let ([sum 0])
                  (with-row
                   (lambda (row)
                     (set! sum (+ sum (first row)))))
                  
                  (check-equal? sum correct_sum)))))
           (lambda ()
             (delete-file "test2.xlsx"))))))))

(run-tests test-test2)