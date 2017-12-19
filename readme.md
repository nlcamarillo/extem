# Extem

An XLSX template engine. 

The idea is to create the desired layout in Excel, possibly by your client. This would include all sorts of styling, row heights, images, etc.

The supply your template and your data to Extem and create the resulting XSLX.

## Install

    npm install extem

## Usage

    let { Extem } = require('extem');

    let myData = {

    };

    Extem.read('./template.xlsx').then((template) => {
        template.evaluate(myData);

        return template.write('./generated.xlsx').then(() => {
            console.log('written');
        })
    }).catch(console.log);

## Creating templates

Extem works with [jsonata](http://jsonata.org/) selectors. These selectors can be applied to cells, but also to ranges. When applied to ranges, every selector inside that range is processed in the scope of the range selector.

Suppose you have the following structure

    {
        list: [1,2,3],
        data: 'foo',
        map: {
            qux: 'moo'
        }
    }

A cell in the sheet with the value `${data}` would result `foo` being rendered in that cell

A cell in the sheet with the value `|{list}` would render 3 cells with `1`, `2` and `3` vertically. All cells below the template cell are pushed down. Similarly, `_{list}` would render a horizontal list, pushing cells to the right (in ltr excel). 

To define a range scope create a cell somewhere and fill it with the following formula

    =IFERROR(N(A3:C5), "${map}")

The syntax is a bit weird (we are looking for a better solution), but has the following benefits:

- it renders the template text in the cell, as N of a range always results in an error
- the range it points to is inspectable and eaily adjustable. Also, with "trace precedents", you get nice insight in the scope structure of your document.

Any cell (or other range) inside that range is now in the scope of `map`, so, the a cell with `${qux}` would render `moo`.
