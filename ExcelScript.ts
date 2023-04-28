namespace ExcelScript {
    export class Worksheet{
        constructor(){}
        public getUsedRange() :any{}
        public  getRange(cell : any) : any{}

        

    };
    export class Workbook{
        constructor(){}
        public getWorksheet(arg : any): any{}
        public getFirstWorksheet(): any{}
        public addWorksheet(arg : any){}s

    }
    export class Range{
        constructor(){}
        public getValues() : any{}
        public getRowCount() : any{}

    }
}