type ResultOk<T> = {
    ok: true;
    data: T;
};

type ResultError<T> = {
    ok: false;
    error: T;
};

export type Result<T, E = Error | string | undefined> =
    | ResultOk<T>
    | ResultError<E>;
