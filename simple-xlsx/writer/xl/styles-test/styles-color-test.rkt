#lang racket

(require rackunit/text-ui)

(require "../../../lib/lib.rkt")

(require rackunit "../styles.rkt")

(define test-styles
  (test-suite
   "test-styles"

   (test-case
    "test-empty-style"
    
    (call-with-input-file "empty-test.dat"
      (lambda (expected)
        (call-with-input-string
         (write-styles '("FF0000" "00FF00" "0000FF"))
         (lambda (actual)
           (check-lines? expected actual))))))

   (test-case
    "test-color-style"
    
    (call-with-input-file "color-test.dat"
      (lambda (expected)
        (call-with-input-string
         (write-styles '("FF0000" "00FF00" "0000FF"))
         (lambda (actual)
           (check-lines? expected actual)))))
   )))

(run-tests test-styles)
