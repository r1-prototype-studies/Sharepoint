export class NewCustomLibraryLibrary {
  public name(): string {
    return "NewCustomLibraryLibrary";
  }

  public getCurrentTime(): string {
    let currentDate: Date;
    let str: string;

    currentDate = new Date();

    str = "<br> Today's Date is : " + currentDate.toDateString();
    str += "<br> Current TIme is : " + currentDate.toTimeString();

    return str;
  }
}
