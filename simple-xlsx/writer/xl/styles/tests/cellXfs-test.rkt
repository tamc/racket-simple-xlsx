#lang racket

(require rackunit/text-ui)

(require "../../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-cellXfs"
    
    (let ([style_list (list #hash((fill . 1)) #hash((fill . 2)) #hash((fill . 3)))])
      
      (call-with-input-file "cellXfs-test.dat"
        (lambda (expected)
          (call-with-input-string
           (write-cellXfs style_list)
           (lambda (actual)
             (check-lines? expected actual)))))
      ))))

(run-tests test-styles)
