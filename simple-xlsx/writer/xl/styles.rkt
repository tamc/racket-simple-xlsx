#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../lib/lib.rkt")
(require "../../xlsx.rkt")

(provide (contract-out
          [write-styles (-> list? list? string?)]
          [write-styles-file (-> path-string? (is-a?/c xlsx%) void?)]
          [write-header (-> string?)]
          [write-fonts (-> string?)]
          [write-fills (-> list? string?)]
          [write-borders (-> string?)]
          [write-cellStyleXfs (-> string?)]
          [write-cellXfs (-> list? string?)]
          [write-cellStyles (-> string?)]
          [write-dxfs (-> string?)]
          [write-footer (-> string?)]
          ))

(define S string-append)

(define (write-header) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
})

(define (write-fonts) @S{
<fonts count="1">
  <font>
    <sz val="11"/>
    <color theme="1"/>
    <name val="宋体"/>
    <family val="2"/>
    <charset val="134"/>
    <scheme val="minor"/>
  </font>
</fonts>
})

(define (write-fills fill_list) @S{
<fills count="@|(number->string (add1 (length fill_list)))|">
  <fill><patternFill patternType="none"/></fill>
@|(let loop ([loop_list fill_list]
             [result_str ""])
    (if (not (null? loop_list))
      (let ([fgColor (hash-ref (car loop_list) 'fgColor "FFFFFF")])
        (loop 
          (cdr loop_list)
          (string-append result_str (format "  <fill><patternFill patternType=\"solid\"><fgColor rgb=\"~a\"/><bgColor indexed=\"64\"/></patternFill></fill>\n" fgColor))))
        result_str))|</fills>
})

(define (write-borders) @S{
<borders count="1">
  <border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
})

(define (write-cellStyleXfs) @S{
<cellStyleXfs count="1">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf>
</cellStyleXfs>
})

(define (write-cellXfs style_list) @S{
<cellXfs count="@|(number->string (add1 (length style_list)))|">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf>
@|(let loop ([loop_list style_list]
             [result_str ""])
    (if (not (null? loop_list))
      (let ([fill (hash-ref (car loop_list) 'fill 0)])
        (loop
          (cdr loop_list)
          (string-append result_str (format "  <xf numFmtId=\"0\" fontId=\"0\" fillId=\"~a\" borderId=\"0\" xfId=\"0\"><alignment vertical=\"center\"/></xf>\n" fill))))
        result_str))|</cellXfs>
})

(define (write-cellStyles) @S{
<cellStyles count="1"><cellStyle name="常规" xfId="0" builtinId="0"/></cellStyles>
})

(define (write-dxfs) @S{
<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
})

(define (write-footer) @S{
</styleSheet>
})

(define (write-styles style_list fill_list) @S{
@|(write-header)|

@|(write-fonts)|

@|(write-fills fill_list)|

@|(write-borders)|

@|(write-cellStyleXfs)|

@|(write-cellXfs style_list)|

@|(write-cellStyles)|

@|(write-dxfs)|

@|(write-footer)|
})

(define (write-styles-file dir xlsx)
  (make-directory* dir)

  (with-output-to-file (build-path dir "styles.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-styles (send xlsx get-color-list))))))
