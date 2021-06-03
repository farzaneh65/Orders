import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import {IPeoplePickerState} from './IPeoplePickerState'
import {IOrdersList} from './IOrdersList'

export interface IOrdersState{
    status: string,
    orderListItem: IOrdersList,
    filePickerResult:IFilePickerResult[],
    Users:string[]
    message:string
}