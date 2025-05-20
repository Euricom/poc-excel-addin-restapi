export declare enum PropertyType {
    Handelsgelijkvloers = "Handelsgelijkvloers",// Commercial ground floor
    Studios = "Studios",
    Appartementen1Slpk = "Appartementen1Slpk",// 1-bedroom apartments
    Appartementen2Slpk = "Appartementen2Slpk",// 2-bedroom apartments
    Appartementen3Slpk = "Appartementen3Slpk",// 3-bedroom apartments
    Kelderbergingen = "Kelderbergingen",// Basement storage
    Yogaruimte = "Yogaruimte"
}
export interface PropertyRentItem {
    id: string;
    type: PropertyType;
    areaM2: number;
    pricePerM2: number;
    totalPricePerYear: number;
    totalPricePerMonth: number;
    location?: string;
    occupancyRate?: number;
    lastUpdated?: string;
}
export interface TotalRent {
    id: string;
    totalAreaM2: number;
    totalPricePerYear: number;
    totalPricePerMonth: number;
    averagePricePerM2?: number;
    calculationDate: string;
}
export interface PropertyPortfolio {
    id: string;
    name: string;
    description?: string;
    location?: string;
    propertyItems: PropertyRentItem[];
    totalRent: TotalRent;
    createdAt: string;
    updatedAt: string;
}
export interface User {
    id: string;
    username: string;
    name: string;
    role: "admin" | "analyst" | "viewer";
}
export interface AuthResponse {
    token: string;
    user: User;
}
export interface BulkUpdateRequest {
    portfolioId: string;
    changes: {
        id: string;
        field: string;
        value: any;
    }[];
}
export interface BulkUpdateResponse {
    results: {
        id: string;
        field: string;
        status: "success" | "error";
        message?: string;
        value?: any;
    }[];
}
