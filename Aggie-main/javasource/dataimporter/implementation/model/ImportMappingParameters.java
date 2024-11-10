package dataimporter.implementation.model;

import com.mendix.integration.ActionWhenNoObjectFound;
import com.mendix.integration.ShouldCommit;

public class ImportMappingParameters {

    private final String importMappingName;
    private final ActionWhenNoObjectFound actionWhenNoObjectFound;
    private final int limit;
    private final ShouldCommit shouldCommit;
    private final String sheetName;

    public ImportMappingParameters(
            String importMappingName,
            ActionWhenNoObjectFound actionWhenNoObjectFound,
            int limit,
            ShouldCommit shouldCommit,
            String sheetName) {

        this.importMappingName = importMappingName;
        this.actionWhenNoObjectFound = actionWhenNoObjectFound;
        this.limit = limit;
        this.shouldCommit = shouldCommit;
        this.sheetName = sheetName;
    }

    public String getImportMappingName() {
        return importMappingName;
    }

    public ActionWhenNoObjectFound getActionWhenNoObjectFound() {
        return actionWhenNoObjectFound;
    }

    public int getLimit() {
        return limit;
    }

    public ShouldCommit getShouldCommit() {
        return shouldCommit;
    }


    public String getSheetName() {
        return sheetName;
    }

    @Override
    public String toString() {
        return "ImportMappingParameters{" +
                ", importMappingName= '" + importMappingName + '\'' +
                ", actionWhenNoObjectFound= " + actionWhenNoObjectFound +
                ", limit= " + limit +
                ", shouldCommit= " + shouldCommit +
                ", sheetName= " + sheetName +
                '}';
    }
}
