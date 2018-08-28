#lang racket

(require rackunit/text-ui)

(require rackunit "../styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-styles"

    (check-equal? 
     (write-styles '("FF0000" "00FF00" "0000FF"))

     (call-with-input-file "color-test.dat"
       (lambda (in) (port->string in)))))
   ))

(run-tests test-styles)
