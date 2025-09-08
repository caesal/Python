def twoSum( nums, target):
    all_result = []
    for x in range(len(nums)):
        for y in range(x+1,len(nums)):
            z = nums[x] + nums[y]
            if z == target :
                all_result = [x,y]
                print(all_result)
                return all_result

twoSum([2,7,11,15], 9)
