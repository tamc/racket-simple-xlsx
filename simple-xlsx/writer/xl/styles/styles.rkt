#lang at-exp racket/base

(require racket/port)
(require racket/file)
(require racket/class)
(require racket/list)
(require racket/contract)

(require "../../../lib/lib.rkt")

(provide (contract-out
          [write-styles (-> list? list? list? string?)]
          [write-header (-> string?)]
          [write-fonts (-> list? string?)]
          [write-fills (-> list? string?)]
          [write-borders (-> string?)]
          [write-cellStyleXfs (-> string?)]
          [write-cellXfs (-> list? string?)]
          [write-cellStyles (-> string?)]
          [write-dxfs (-> string?)]
          [write-footer (-> string?)]
          [write-styles-file (-> path-string? list? list? list? void?)]
          ))

(define S string-append)

(define (write-header) @S{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>

<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
})

(define (write-fonts font_list) @S{
<fonts count="@|(number->string (add1 (length font_list)))|">
  <font>
    <sz val="11"/>
    <color theme="1"/>
    <name val="宋体"/>
    <family val="2"/>
    <charset val="134"/>
    <scheme val="minor"/>
  </font>
@|(with-output-to-string
    (lambda ()
      (let loop ([loop_list font_list])
        (when (not (null? loop_list))
          (printf "\n")
          (let ([fontSize (hash-ref (car loop_list) 'fontSize 11)])
            (printf "  <font>\n")
            (printf "    <sz val=\"~a\"/>\n" fontSize)
            (printf "    <color theme=\"1\"/>\n")
            (printf "    <name val=\"宋体\"/>\n")
            (printf "    <family val=\"2\"/>\n")
            (printf "    <charset val=\"134\"/>\n")
            (printf "    <scheme val=\"minor\"/>\n")
            (printf "  </font>\n")
            (loop (cdr loop_list)))))))|</fonts>
})

(define (write-fills fill_list) @S{
<fills count="@|(number->string (add1 (length fill_list)))|">
  <fill><patternFill patternType="none"/></fill>
@|(let loop ([loop_list fill_list]
             [result_str ""])
    (if (not (null? loop_list))
      (let ([backgroundColor (hash-ref (car loop_list) 'fgColor "FFFFFF")])
        (loop 
          (cdr loop_list)
          (string-append result_str (format "  <fill><patternFill patternType=\"solid\"><fgColor rgb=\"~a\"/><bgColor indexed=\"64\"/></patternFill></fill>\n" backgroundColor))))
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
@|(with-output-to-string
    (lambda ()
      (let loop ([loop_list style_list])
        (when (not (null? loop_list))
          (let ([fill (hash-ref (car loop_list) 'fill 0)]
                [font (hash-ref (car loop_list) 'font 0)])
            (printf "  <xf numFmtId=\"0\" fontId=\"~a\" fillId=\"~a\" borderId=\"0\" xfId=\"0\"" font fill)
            (when (not (= font 0)) (printf " applyFont=\"1\""))
            (when (not (= fill 0)) (printf " applyFill=\"1\""))
            (printf "><alignment vertical=\"center\"/></xf>\n"))
          (loop (cdr loop_list))))))|</cellXfs>
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

(define (write-styles style_list fill_list font_list) @S{
@|(write-header)|

@|(prefix-each-line (write-fonts font_list) "  ")|

@|(prefix-each-line (write-fills fill_list) "  ")|

@|(prefix-each-line (write-borders) "  ")|

@|(prefix-each-line (write-cellStyleXfs) "  ")|

@|(prefix-each-line (write-cellXfs style_list) "  ")|

@|(prefix-each-line (write-cellStyles) "  ")|

@|(prefix-each-line (write-dxfs) "  ")|

@|(write-footer)|
})

(define (write-styles-file dir style_list fill_list font_list)
  (make-directory* dir)

  (with-output-to-file (build-path dir "styles.xml")
    #:exists 'replace
    (lambda ()
      (printf "~a" (write-styles 
                    style_list
                    fill_list
                    font_list
                    )))))
