import {IPeoplePickerState} from './IPeoplePickerState'

export interface IOrdersList{
    Id: number,
    Title: string,
    Quantity: number,
    DestinationId: number,
    Description: string,
    LinkToFile:string,
   Owners:number[]
}