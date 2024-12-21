use std::ops::Index;

pub trait TryIterator {
    type Item;
    type Error;

    fn try_next(&mut self) -> Result<Option<Self::Item>, Self::Error>;
}

pub struct TryIter<T: TryIterator> {
    inner: T,
    done: bool,
}

impl<T: TryIterator> Iterator for TryIter<T> {
    type Item = Result<T::Item, T::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        match self.inner.try_next() {
            Ok(Some(item)) => Some(Ok(item)),
            Ok(None) => {
                self.done = true;
                None
            }
            Err(e) => {
                self.done = true;
                Some(Err(e))
            }
        }
    }
}

pub trait TryCacheIterator {
    type Item;
    type Error;
    type Cache: AsRef<[Self::Item]> + Index<usize, Output = Self::Item>;

    fn try_cache(&mut self) -> Result<Self::Cache, Self::Error>;
}

pub struct TryCacheIter<T: TryCacheIterator> {
    inner: T,
    cache: T::Cache,
    index: usize,
    done: bool,
}

impl<T: TryCacheIterator> Iterator for TryCacheIter<T> {
    type Item = Result<T::Item, T::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        if self.index < self.cache.as_ref().len() {
            let item = unsafe {
                core::ptr::replace(
                    &self.cache[self.index] as *const _ as *mut _,
                    core::mem::zeroed(),
                )
            };
            self.index += 1;
            return Some(Ok(item));
        }

        match self.inner.try_cache() {
            Ok(cache) => {
                self.cache = cache;
                self.index = 0;
                if self.cache.as_ref().is_empty() {
                    self.done = true;
                    None
                } else {
                    self.next()
                }
            }
            Err(e) => {
                self.done = true;
                Some(Err(e))
            }
        }
    }
}
