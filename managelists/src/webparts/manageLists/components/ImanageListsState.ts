import { 
    IList,
    ISPSite
 } from './../common/IObjects';  
interface IManageListsState{  
    lists: IList[];  
    rootSite: ISPSite;
}  
export default IManageListsState;  